VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm105Bypass 
   BackColor       =   &H00DBE6E6&
   Caption         =   "POC & Bypass"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis105.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   11400
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   1335
      TabIndex        =   65
      Top             =   45
      Width           =   13110
      _ExtentX        =   23125
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
      Caption         =   "환자 기본정보"
   End
   Begin VB.Frame fraPass 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   2460
      TabIndex        =   55
      Top             =   1605
      Width           =   2730
      Begin VB.CommandButton cmdApplyBypass 
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용"
         Height          =   330
         Left            =   1770
         Style           =   1  '그래픽
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2295
         Width           =   915
      End
      Begin VB.CommandButton cmdCloseBypass 
         BackColor       =   &H00DBE6E6&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2430
         Style           =   1  '그래픽
         TabIndex        =   57
         Top             =   15
         Width           =   300
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1965
         Left            =   0
         TabIndex        =   56
         Top             =   315
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3466
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "검사항목"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "검체"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "검사코드"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "검체코드"
            Object.Width           =   2540
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   285
         Index           =   0
         Left            =   15
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   15
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "♣ ByPass 항목리스트"
      End
   End
   Begin VB.CommandButton cmdByPass 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Bypass 항목"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3945
      Style           =   1  '그래픽
      TabIndex        =   54
      Top             =   1275
      Width           =   1245
   End
   Begin VB.Frame fraAcc 
      BackColor       =   &H00EAE7E3&
      BorderStyle     =   0  '없음
      Height          =   2250
      Left            =   11325
      TabIndex        =   50
      Top             =   1575
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton cmdCloseFra 
         BackColor       =   &H00DBE6E6&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2850
         Style           =   1  '그래픽
         TabIndex        =   52
         Top             =   15
         Width           =   255
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   240
         Left            =   30
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   15
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   423
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "♣ 접수번호리스트"
      End
      Begin MSComctlLib.ListView lvwCount 
         Height          =   2010
         Left            =   15
         TabIndex        =   53
         Top             =   240
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   3545
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "순번"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "접수번호"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "확인"
            Object.Width           =   1058
         EndProperty
      End
   End
   Begin VB.ListBox lstTestList 
      Appearance      =   0  '평면
      BackColor       =   &H00F7F3F8&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   555
      Sorted          =   -1  'True
      TabIndex        =   48
      Top             =   2370
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.ListBox lstSpcList 
      Appearance      =   0  '평면
      BackColor       =   &H00FCE9F7&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   3315
      TabIndex        =   47
      Top             =   2370
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdAcc 
      BackColor       =   &H00DBE6E6&
      Caption         =   "..."
      Height          =   300
      Left            =   14130
      Style           =   1  '그래픽
      TabIndex        =   46
      Top             =   1260
      Width           =   315
   End
   Begin VB.CommandButton cmdNextData 
      BackColor       =   &H00EAE7E3&
      Caption         =   "<< (&P)"
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
      Index           =   0
      Left            =   10260
      Style           =   1  '그래픽
      TabIndex        =   45
      Top             =   1260
      Width           =   705
   End
   Begin VB.CommandButton cmdNextData 
      BackColor       =   &H00EAE7E3&
      Caption         =   "(&N) >>"
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
      Index           =   1
      Left            =   10980
      Style           =   1  '그래픽
      TabIndex        =   44
      Top             =   1260
      Width           =   705
   End
   Begin VB.Frame fraAccNo 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   300
      Left            =   7725
      TabIndex        =   42
      Top             =   1275
      Width           =   2535
      Begin MSMask.MaskEdBox mskAccNo 
         Height          =   300
         Left            =   945
         TabIndex        =   43
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "&&-######-####"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   4
         Left            =   15
         TabIndex        =   66
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdAccession 
      BackColor       =   &H00EAE7E3&
      Caption         =   "＞＞접수"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   5220
      Style           =   1  '그래픽
      TabIndex        =   40
      Tag             =   "19907"
      Top             =   3075
      Width           =   510
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
      Height          =   1965
      Left            =   6960
      TabIndex        =   37
      Tag             =   "20002"
      Top             =   6555
      Width           =   7500
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
         Left            =   6990
         Picture         =   "Lis105.frx":000C
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   1575
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1620
         Left            =   75
         TabIndex        =   39
         Top             =   270
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   2858
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis105.frx":053E
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
      Height          =   1965
      Left            =   75
      TabIndex        =   31
      Tag             =   "20003"
      Top             =   6555
      Width           =   6885
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
         Left            =   6480
         Picture         =   "Lis105.frx":05DB
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   945
         Width           =   315
      End
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
         Left            =   6480
         Picture         =   "Lis105.frx":0B0D
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   1575
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   990
         Left            =   90
         TabIndex        =   34
         Top             =   270
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1746
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis105.frx":103F
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
         TabIndex        =   35
         Top             =   1530
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis105.frx":10DC
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
         TabIndex        =   36
         Top             =   1260
         Width           =   1545
      End
   End
   Begin VB.PictureBox picRst 
      BackColor       =   &H00E0E0E0&
      Height          =   4935
      Left            =   5760
      ScaleHeight     =   4875
      ScaleWidth      =   8655
      TabIndex        =   26
      Top             =   1590
      Width           =   8715
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   255
         Left            =   0
         TabIndex        =   27
         ToolTipText     =   "자료를 가져오고 있읍니다."
         Top             =   4620
         Visible         =   0   'False
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   4800
         Left            =   30
         TabIndex        =   28
         Tag             =   "20001"
         Top             =   0
         Width           =   8535
         _Version        =   196608
         _ExtentX        =   15055
         _ExtentY        =   8467
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
         MaxCols         =   18
         MaxRows         =   0
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis105.frx":1179
         VisibleCols     =   10
         TextTip         =   2
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
         TabIndex        =   29
         Top             =   2520
         Width           =   6675
      End
   End
   Begin VB.Frame fraPtInfo 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1335
      TabIndex        =   3
      Tag             =   "104"
      Top             =   285
      Width           =   13125
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   9
         Left            =   7530
         TabIndex        =   74
         Top             =   555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   8
         Left            =   7530
         TabIndex        =   73
         Top             =   195
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   6
         Left            =   2310
         TabIndex        =   71
         Top             =   555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   5
         Left            =   2310
         TabIndex        =   70
         Top             =   195
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "처방 구분"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00EAE7E3&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   10875
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   24
         Tag             =   "WardId"
         Top             =   540
         Width           =   315
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00EAE7E3&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   9645
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   23
         Tag             =   "DoctId"
         Top             =   180
         Width           =   285
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00EAE7E3&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   9645
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   22
         Tag             =   "DeptCd"
         Top             =   555
         Width           =   285
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
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
         Left            =   990
         MaxLength       =   10
         TabIndex        =   15
         Top             =   195
         Width           =   1305
      End
      Begin VB.TextBox txtDoctorId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
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
         Left            =   8475
         TabIndex        =   14
         Top             =   195
         Width           =   1170
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3255
         ScaleHeight     =   285
         ScaleWidth      =   4230
         TabIndex        =   10
         Top             =   195
         Width           =   4260
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외래처방"
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
            Left            =   165
            TabIndex        =   13
            Tag             =   "10108"
            Top             =   30
            Width           =   1095
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "병동처방"
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
            Index           =   1
            Left            =   1485
            TabIndex        =   12
            Tag             =   "10109"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "응급실처방"
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
            Height          =   240
            Index           =   2
            Left            =   2775
            TabIndex        =   11
            Tag             =   "10110"
            Top             =   30
            Width           =   1290
         End
      End
      Begin VB.TextBox txtDeptCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
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
         Left            =   8475
         TabIndex        =   7
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtWardId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11265
         TabIndex        =   6
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txtRoomId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12000
         TabIndex        =   5
         Top             =   615
         Width           =   525
      End
      Begin VB.TextBox txtBedId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12660
         TabIndex        =   4
         Top             =   615
         Width           =   390
      End
      Begin MedControls1.LisLabel lblDob 
         Height          =   300
         Left            =   6255
         TabIndex        =   8
         Top             =   555
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         BackColor       =   15857140
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   990
         TabIndex        =   9
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         BackColor       =   15857140
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
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   330
         Left            =   9945
         TabIndex        =   16
         Top             =   180
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   150
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   68
         Top             =   195
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "환자 ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   3
         Left            =   60
         TabIndex        =   69
         Top             =   555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "성    명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   7
         Left            =   5310
         TabIndex        =   72
         Top             =   555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   10
         Left            =   9945
         TabIndex        =   75
         Top             =   555
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin VB.Label lblAgeDiv 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4575
         TabIndex        =   20
         Top             =   615
         Width           =   405
      End
      Begin VB.Label lblAge 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4230
         TabIndex        =   19
         Top             =   615
         Width           =   345
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3300
         TabIndex        =   18
         Top             =   615
         Width           =   690
      End
      Begin VB.Label lblWard 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "           -          -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11205
         TabIndex        =   17
         Tag             =   "107"
         Top             =   555
         Width           =   1860
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BorderStyle     =   1  '단일 고정
         Caption         =   "             /"
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
         Height          =   315
         Left            =   3255
         TabIndex        =   21
         Top             =   555
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료 (&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장 (&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   5760
      TabIndex        =   30
      Top             =   1275
      Width           =   1965
      _ExtentX        =   3466
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
      Caption         =   "접수번호별 결과등록"
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   75
      TabIndex        =   25
      Top             =   1275
      Width           =   3900
      _ExtentX        =   6879
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
      Caption         =   "처방등록"
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   4860
      Left            =   75
      TabIndex        =   49
      Tag             =   "10114"
      Top             =   1605
      Width           =   5130
      _Version        =   196608
      _ExtentX        =   9049
      _ExtentY        =   8572
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      FormulaSync     =   0   'False
      MaxCols         =   23
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis105.frx":1829
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin FPSpread.vaSpread tblCollect 
      Height          =   4860
      Left            =   60
      TabIndex        =   63
      Tag             =   "10114"
      Top             =   1605
      Visible         =   0   'False
      Width           =   5130
      _Version        =   196608
      _ExtentX        =   9049
      _ExtentY        =   8573
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      GridColor       =   14737632
      MaxCols         =   36
      MaxRows         =   19
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "Lis105.frx":44A6
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   19
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   75
      TabIndex        =   64
      Top             =   45
      Width           =   1245
      _ExtentX        =   2196
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
      Caption         =   "업무구분"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   75
      TabIndex        =   60
      Top             =   270
      Width           =   1260
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ByPass"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   62
         Top             =   540
         Width           =   1035
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "POC"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   61
         Top             =   255
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   285
      Index           =   1
      Left            =   12705
      TabIndex        =   67
      Top             =   1275
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
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
      Caption         =   "접수대상확인"
      Appearance      =   0
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
      TabIndex        =   41
      Top             =   8715
      Width           =   2385
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8640
      Width           =   9675
   End
End
Attribute VB_Name = "frm105Bypass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsTemplete  As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
'Private WithEvents objMyList    As clspopuplist
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private Const TAG_DEPT& = 1
Private Const TAG_WARD& = 2
Private Const TAG_DOCT& = 3
'Private WithEvents objCodeList  As clspopuplist
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
'Private WithEvents mnuPopup     As menu
'Private WithEvents mnuDelete    As menu

Private Const MENU_DELETE& = 1
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1

Private objSqlStmt  As New clsLISSqlStatement     ' SQL 클래스
Private objPatient  As New clsPatient
Private objOrder    As New clsLISOrder
Private objCollect  As New clsLISCollectioin
Private objAccess   As New clsLISAccession
Private objPtInfo   As New clsPatientInfo


Private objDic      As clsDictionary

Private blnClearFg  As Boolean

Private gintTemplete    As Integer
Private ClearFg         As Boolean
Private MsgFg           As Boolean

Private Sub cmdApplyBypass_Click()
    Dim iTmx As ListItem
    For Each iTmx In lvw.ListItems
        If iTmx.Checked = True Then
            Call ByPassMove(iTmx.SubItems(2), iTmx.SubItems(3))
        End If
    Next
    fraPass.Visible = False
End Sub

Private Sub cmdByPass_Click()
    Dim RS      As Recordset
    Dim SSQL    As String
    
    SSQL = " SELECT a.cdval1 as testcd,a.field1 as spccd,b.field5 as spcnm,c.abbrnm10 as testnm " & _
           " FROM " & _
                    T_LAB032 & " a," & T_LAB032 & " b," & T_LAB001 & " c" & _
           " WHERE " & _
                     DBW("a.cdindex=", LC3_ByPass) & _
           " AND " & DBW("b.cdindex=", LC3_Specimen) & _
           " AND a.field1=b.cdval1" & _
           " AND a.cdval1=c.testcd"
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    lvw.ListItems.Clear
    
    If Not RS.EOF Then
        Dim iTmx As ListItem
        
        Do Until RS.EOF
            Set iTmx = lvw.ListItems.Add(, , RS.Fields("testnm").Value & "")
            iTmx.SubItems(1) = RS.Fields("spcnm").Value & ""
            iTmx.SubItems(2) = RS.Fields("testcd").Value & ""
            iTmx.SubItems(3) = RS.Fields("spccd").Value & ""
            RS.MoveNext
        Loop
        fraPass.Visible = True
    Else
        MsgBox "ByPass항목으로 지정된 검사가 없습니다.", vbInformation + vbOKOnly, "Info"
    End If
    Set RS = Nothing
End Sub


Public Sub ByPassMove(ByVal sTestcd As String, ByVal sSpcCd As String)
    Dim tmpTestCd   As String
    Dim tmpSpcCd    As String
    Dim ii          As Integer
    Dim objSQL As clsLISSqlStatement
    Dim RS As Recordset
    
    Set objSQL = New clsLISSqlStatement
    Set RS = New Recordset
    
    RS.Open objSQL.GetItemInfo(sTestcd), DBConn
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enORDSHEET.tcTESTCD: tmpTestCd = Trim(.Value)
            .Col = enORDSHEET.tcSPCCD: tmpSpcCd = Trim(.Value)
            If sTestcd = tmpTestCd And tmpSpcCd = sSpcCd Then Exit Sub
        Next
        
        If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
        .Row = .DataRowCnt + 1
'        Call rs.KeyChange(sTestcd)
        
        .Col = enORDSHEET.tcTESTNM: .Value = RS.Fields("testnm").Value & ""  ' 처방명
        .Col = enORDSHEET.tcTESTCD: .Value = RS.Fields("testcd").Value & "" ' 처방코드
        
        
        
        .Col = enORDSHEET.tcINSURFG:    .Value = RS.Fields("insurfg").Value & ""   ' 급여구분
        .Col = enORDSHEET.tcSPCCD:      .Value = sSpcCd     ' 검체코드
        .Col = enORDSHEET.tcREQDTTM:    .Value = Format(GetSystemDate, "YYYY/MM/DD HH:MM").Value & ""
        .Col = enORDSHEET.tcSTATFG:     .Value = RS.Fields("statfg").Value & ""    ' **응급여부(해당건물)
    '***건물정보 사용
        If P_ApplyBuildingInfo Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcSTATCHK: .CellType = 10     'CellTypeCheckBox
                                             .TypeCheckCenter = True
            Else
                .Col = enORDSHEET.tcSTATCHK: .CellType = 5  'CellTypeStaticText
            End If
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = 10     'CellTypeCheckBox
                                             .TypeCheckCenter = True
        End If
        .Col = enORDSHEET.tcWORKAREA:   .Value = RS.Fields("workarea").Value & ""  ' WorkArea
        .Col = enORDSHEET.tcSTORECD:    .Value = RS.Fields("storecd").Value & ""   ' 보관구분
        .Col = enORDSHEET.tcRNDFG:      .Value = RS.Fields("rndfg").Value & ""     ' 아침채혈 여부
        .Col = enORDSHEET.tcTESTDIV:    .Value = RS.Fields("testdiv").Value & ""   ' 테스트구분(0:일반,1:기타,2:미생물)
        .Col = enORDSHEET.tcMULTIFG:    .Value = RS.Fields("multifg").Value & ""   ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:     .Value = RS.Fields("spcgrp").Value & ""    ' 검체군
        .Col = enORDSHEET.tcABBRNM:     .Value = RS.Fields("abbrnm5").Value & ""   ' 약어명
        .Col = enORDSHEET.tcBARCNT:     .Value = RS.Fields("barcnt").Value & ""    ' 라벨출력장수
        .Col = enORDSHEET.tcTESTFLAG:   .Value = RS.Fields("testfg").Value & ""    ' **검사가능여부(해당건물)
        
    '***건물정보 사용
        If P_ApplyBuildingInfo Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd      ' ** 해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab      ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
        Else
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd      ' ** 해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = LABName
        End If
        .Col = enORDSHEET.tcSPCABBR:    .Value = RS.Fields("spcnm").Value & ""     ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:     .Value = RS.Fields("labdiv").Value & ""    ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:   .Value = RS.Fields("labrange").Value & ""  ' 미생물 접수번호 범위
    End With
    
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub


Private Sub cmdClear_Click()
    Call ICSPatientMark
    Call ClearRtn
    blnClearFg = True
    txtPtId.SetFocus
End Sub


Private Sub cmdCloseBypass_Click()
    fraPass.Visible = False
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Set objMyList = New clspopuplist
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        Select Case Index
            Case 1
                .FormCaption = "처방의 조회"
                .ColumnHeaderText = "처방의ID;처방의명"
                .Tag = TAG_DOCT
                .LoadPopUp GetSQLDoctList
                
'                 .Caption = "처방의 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "처방의ID, 처방의명"
'                 Call .ListPop(GETSQLDOCT, 1640, cmdHelpList(Index).Left)
            Case 2
                .FormCaption = "진료과 조회"
                .ColumnHeaderText = "진료과코드;진료과명"
                .Tag = TAG_DEPT
                .LoadPopUp GetSQLDeptList
                
'                 .Caption = "진료과 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "진료과코드, 진료과명"
'                 Call .ListPop(GETSQLDEPT, 1640, cmdHelpList(Index).Left) ', objLisComCode.DeptCd)
'                 If txtPtId <> "" Then
'                 End If
            Case 3
                .FormCaption = "병동 조회"
                .ColumnHeaderText = "병동코드;병동명"
                .Tag = TAG_WARD
                .LoadPopUp GetSQLWardList
                
'                 .Caption = "병동 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "병동코드,병동명"
'                 Call .ListPop(GETSQLWARD, 1640, 10550) ', objLisComCode.WardId)
        End Select
    End With

    Set objMyList = Nothing
'    Set objData = Nothing

End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub


'% 폼 로드...
Private Sub Form_Load()

    Dim tmpDate As Date
    Dim i As Integer
'    Dim objProgress As New clsProgress
    Dim objProgress As New clsProgress

    Me.Show
    Call ClearRtn

    DoEvents
    
    With objProgress
        .Container = MainFrm.stsBar
        .Message = "검사항목 리스트를 로드하고 있습니다..."
    End With
    
'    objProgress.CaptionOn = False
'    objProgress.MSG = "검사항목 리스트를 로드하고 있습니다..."
'    objProgress.mode = 0
'    objProgress.Visible = True
'    objProgress.Value = 0

    '다빈도 항목 / 검사항목 로드...

    MouseRunning

    Call ItemList(lstTestList, objProgress)
    objProgress.Value = objProgress.Max
'    objProgress.Visible = False
    Set objProgress = Nothing

    MouseDefault
    
    txtPtId.SetFocus

End Sub

Private Sub ItemList(ByRef lstList As Object, Optional ByRef barStatus As Variant)

'    Dim tmpTestCd As String
'    Dim tmpTestNm As String
'
'    If Not IsMissing(barStatus) Then barStatus.Max = objLisItem.RecordCount + 1
'
'    DoEvents
'
'    With lstList
'        .Clear
'        medLockWindowUpdate (.hwnd)
'        objLisItem.MoveFirst
'
'        While (Not objLisItem.EOF)
'            tmpTestNm = Mid(objLisItem.Fields("TestNm"), 1, 40)
'            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' 검사명
'            tmpTestCd = Trim(Mid(objLisItem.Fields("TestCd"), 1, 9))
'            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' 검사코드
'
'            If Trim(tmpTestCd) <> "" And objLisItem.Fields("testdiv") = "0" Then
'                .AddItem tmpTestNm & tmpTestCd & "1"  '검사명기준
'                .AddItem tmpTestCd & tmpTestNm & "2"  '검사코드기준
'            End If
'
'            If Not IsMissing(barStatus) Then barStatus.Value = barStatus.Value + 1
'            DoEvents
'            objLisItem.MoveNext
'        Wend
'        .Visible = False
'        medLockWindowUpdate (0&)
'    End With
    
    
    Dim i As Integer
    Dim tmpTestCd As String
    Dim tmpTestNm As String
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT a.testnm, a.abbrnm5, a.testcd, b.spccd, b.statfg, a.workarea, b.storecd, b.rndfg, " & _
            "        b.labelcnt, b.statflags, a.testdiv, c.field1 as MultiFg, c.field2 as SpcGrp, c.field5 as SpcNm, " & _
            "        d.field2 as LabDiv, e.field2 as LabRange, '1' InsurFg " & _
            " FROM " & T_LAB032 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & _
                       T_LAB004 & " b, " & T_LAB001 & " a " & _
            " WHERE  a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
            "                     WHERE testcd = a.testcd ) " & _
            " AND   (a.detailfg = '' or a.detailfg is null) " & _
            " AND    a.testcd = b.testcd " & _
            " AND    b.seq = ( SELECT min(seq) FROM " & T_LAB004 & _
            "                  WHERE testcd = b.testcd ) " & _
            " AND   (b.expdt = '' or b.expdt is null)" & _
            " AND    b.applydt = ( SELECT max(applydt) FROM " & T_LAB004 & _
            "                      WHERE testcd = b.testcd AND spccd = b.spccd AND seq=b.seq) " & _
            " AND    c.cdindex = 'C215' " & _
            " AND    c.cdval1 = b.spccd  " & _
            " AND    d.cdindex = 'C213' " & _
            " AND    d.cdval1 = a.workarea " & _
            " AND    " & DBJ("e.cdindex = 'C217'") & _
            " AND    " & DBJ("e.cdval1 =* c.field2")
    
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    If Not IsMissing(barStatus) Then barStatus.Max = RS.RecordCount + 1
   
    DoEvents
   
    With lstList
        .Clear
        
        medLockWindowUpdate (.hwnd)
        RS.MoveFirst
        
        While (Not RS.EOF)
            
             
            tmpTestNm = Mid(RS.Fields("TestNm").Value & "", 1, 40)
            tmpTestNm = tmpTestNm & Space(40 - Len(tmpTestNm)) & vbTab  ' 검사명
            tmpTestCd = Trim(Mid(RS.Fields("TestCd").Value & "", 1, 9))
            tmpTestCd = tmpTestCd & Space(9 - Len(tmpTestCd)) & vbTab   ' 검사코드
             
            If Trim(tmpTestCd) <> "" Then
            
                tmpStatFg = medGetP(RS.Fields("StatFlags").Value & "", 1, ";") ' 건물별 응급가능 여부
                tmpTestFg = medGetP(RS.Fields("StatFlags").Value & "", 2, ";") ' 건물별 검사가능 여부
         
'                tmpStatFg = Mid(tmpStatFg, gBuildingNo, 1)                   '** 응급여부(해당건물)
'                Rs.Fields("statfg") = tmpStatFg
'                tmpTestFg = Mid(tmpTestFg, gBuildingNo, 1)                   '** 검사가능여부(해당건물)
'                Rs.Fields("testfg") = tmpTestFg
         
                .AddItem tmpTestNm & tmpTestCd & "1"  '검사명기준
                .AddItem tmpTestCd & tmpTestNm & "2"  '검사코드기준
                
            End If
         
            If Not IsMissing(barStatus) Then barStatus.Value = barStatus.Value + 1
         
            DoEvents
            RS.MoveNext
        Wend
        .Visible = False
        medLockWindowUpdate (0&)
        
    End With
    
    Set RS = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '감염관리표시
    Call ICSPatientMark
    Set objDic = Nothing
'    Set mnuPopup = Nothing
    Set objOrder = Nothing
    Set objMyList = Nothing
'    Set mnuDelete = Nothing
    Set objAccess = Nothing
    Set objPtInfo = Nothing
    Set objSqlStmt = Nothing
    Set objPatient = Nothing
    Set objCollect = Nothing
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
End Sub


'% 검체리스트에서 항목 선택을 키보드로 했을 경우...
Private Sub lstSpcList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 32:    'Enter Key 또는 Space
            Call lstSpcList_MouseDown(1, 0, 0, 0)
        Case 27:  'ESC
            lstSpcList.Visible = False
            tblOrdSheet.SetFocus
        Case Else:   '그 밖에...
            tblOrdSheet.SetFocus
            tblOrdSheet.Action = ActionActiveCell
    End Select
End Sub

Private Sub lstSpcList_LostFocus()
    If lstSpcList.Visible Then
        lstSpcList.SetFocus
        Exit Sub
    End If
    tblOrdSheet.SetFocus
End Sub

'% 검체리스트에서 항목 선택을 마우스로 했을 경우...
Private Sub lstSpcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim tmpStr As String

    If Button <> 1 Then Exit Sub

    tmpStr = lstSpcList.List(lstSpcList.ListIndex)

    With tblOrdSheet
        .Col = enORDSHEET.tcSPCCD:       .Value = Trim(medShift(tmpStr, vbTab))       ' 검체코드

        Call medShift(tmpStr, vbTab)

        .Col = enORDSHEET.tcSTATFG:      .Value = Trim(medShift(tmpStr, vbTab))       ' **응급여부(해당건물)
        If .Value = "1" Then
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeStaticText
        End If
        .Col = enORDSHEET.tcSTORECD:     .Value = Trim(medShift(tmpStr, vbTab))       ' 보관구분
        .Col = enORDSHEET.tcMULTIFG:     .Value = Trim(medShift(tmpStr, vbTab))       ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:      .Value = Trim(medShift(tmpStr, vbTab))       ' 검체군
        .Col = enORDSHEET.tcBARCNT:      .Value = Trim(medShift(tmpStr, vbTab))       ' 라벨출력장수
        .Col = enORDSHEET.tcTESTFLAG:    .Value = Trim(medShift(tmpStr, vbTab))       ' **검사가능여부(해당건물)
    '***건물정보 사용
        If ObjSysInfo.UseBuildingInfo = "1" Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd    ' ** 해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab    ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
    '***건물정보 사용하지 않음
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd    ' ** 해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
        End If
        .Col = enORDSHEET.tcSPCABBR:     .Value = Trim(medShift(tmpStr, vbTab))       ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:      .Value = Trim(medShift(tmpStr, vbTab))       ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:    .Value = Trim(medShift(tmpStr, vbTab))       ' 미생물 접수번호 범위

        lstSpcList.Visible = False
        .SetFocus
        .Col = enORDSHEET.tcSTATCHK
        .Action = ActionActiveCell
    End With

End Sub

'% 처방항목 리스트에서 항목 선택을 키보드로 했을 경우...
Private Sub lstTestList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 32:    'Enter Key 또는 Space
            Call lstTestList_MouseDown(1, 0, 0, 0)
        Case 27:  'ESC
            lstTestList.Visible = False
            tblOrdSheet.SetFocus
        Case Else:   '그 밖에...
            tblOrdSheet.SetFocus
            tblOrdSheet.Action = ActionActiveCell
            SendKeys Chr(KeyAscii)
   End Select
End Sub

'% 처방항목 리스트에서 항목 선택을 마우스로 했을 경우...
Private Sub lstTestList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String
    Dim strFlag1 As String, strFlag2 As String
    Dim objSQL As clsLISSqlStatement
    Dim RS As Recordset

    If Button <> 1 Then Exit Sub
    If lstTestList.ListIndex < 0 Then Exit Sub
    
    
    
    tmpStr = lstTestList.List(lstTestList.ListIndex)

    With tblOrdSheet
        tmpField1 = Trim(medShift(tmpStr, vbTab))
        tmpField2 = medShift(tmpStr, vbTab)
        
        
        Set objSQL = New clsLISSqlStatement
        Set RS = New Recordset
        
        If tmpStr = "1" Then
            .Col = enORDSHEET.tcTESTNM:  .Value = Trim(tmpField1)    ' 처방명
            .Col = enORDSHEET.tcTESTCD:  .Value = Trim(tmpField2)    ' 처방코드
'            Call rs.KeyChange(Trim$(tmpField2))
            RS.Open objSQL.GetItemInfo(tmpField2), DBConn
        Else
            .Col = enORDSHEET.tcTESTNM:  .Value = Trim(tmpField2)    ' 처방명
            .Col = enORDSHEET.tcTESTCD:  .Value = Trim(tmpField1)    ' 처방코드
'            Call rs.KeyChange(Trim$(tmpField1))
            RS.Open objSQL.GetItemInfo(tmpField1), DBConn
        End If

        .Col = enORDSHEET.tcINSURFG:     .Value = RS.Fields("insurfg").Value & ""       ' 급여구분
        .Col = enORDSHEET.tcREQDTTM:     '.Value = Format(GetSystemdate & " " & GetSystemdate, _
                                                  CS_DateLongFormat & " " & CS_TimeShortFormat)     ' 희망채취시간
        
                                         .Value = Format(GetSystemDate, CS_DateLongFormat)
                                         .Value = .Value & " " & Format(GetSystemDate, CS_TimeShortFormat)
        .Col = enORDSHEET.tcSPCCD:       .Value = RS.Fields("spccd").Value & ""         ' 검체코드
        .Col = enORDSHEET.tcWORKAREA:    .Value = RS.Fields("workarea").Value & ""      ' WorkArea
        .Col = enORDSHEET.tcSTORECD:     .Value = RS.Fields("storecd").Value & ""       ' StoreCd
        .Col = enORDSHEET.tcRNDFG:       .Value = RS.Fields("rndfg").Value & ""         ' 아침채혈여부
        .Col = enORDSHEET.tcTESTDIV:     .Value = RS.Fields("testdiv").Value & ""       ' 검사구분
        .Col = enORDSHEET.tcMULTIFG:     .Value = RS.Fields("multifg").Value & ""       ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:      .Value = RS.Fields("spcgrp").Value & ""        ' 검체군
        .Col = enORDSHEET.tcABBRNM:      .Value = RS.Fields("abbrnm5").Value & ""       ' 약어명
        .Col = enORDSHEET.tcBARCNT:      .Value = RS.Fields("labelcnt").Value & ""      ' 라벨출력장수
        .Col = enORDSHEET.tcSPCABBR:     .Value = RS.Fields("spcnm").Value & ""         ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:      .Value = RS.Fields("labdiv").Value & ""        ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:    .Value = RS.Fields("labrange").Value & ""      ' 미생물 접수번호 범위
        
        tmpStr = RS.Fields("statflags")
        strFlag1 = medGetP(tmpStr, 1, ";")
        strFlag2 = medGetP(tmpStr, 2, ";")
        RS.Fields("statfg") = Mid(strFlag1, ObjSysInfo.BuildingNo, 1)
        RS.Fields("testfg") = Mid(strFlag2, ObjSysInfo.BuildingNo, 1)
        
        If .Value = "1" Then
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeStaticText
        End If
        .Col = enORDSHEET.tcTESTFLAG:    .Value = RS.Fields("testfg").Value & ""       ' **검사가능여부(해당건물)
    '***건물정보 사용
        If ObjSysInfo.UseBuildingInfo = "1" Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd     ' **해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab                ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
    '***건물정보 사용하지 않음
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd     ' **해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
        End If
        lstTestList.Visible = False
        Call tblOrdSheet_LeaveCell(.Col, .Row, enORDSHEET.tcSPCCD, .Row, False)

    End With
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub mskAccNo_LostFocus()
    Data_Load
End Sub

'Private Sub objCodeList_ListClick(ByVal SelList As String)
'    objPtInfo.RmkCd = medGetP(SelList, 1, vbTab)
'    objPtInfo.RmkNm = medGetP(SelList, 2, vbTab)
'    rtfRemark.Text = objPtInfo.RmkNm
'End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    objPtInfo.RmkCd = objCodeList.SelectedItems(0)
    objPtInfo.RmkNm = objCodeList.SelectedItems(1)
    rtfRemark.Text = objPtInfo.RmkNm
End Sub

'Private Sub objMyList_SendCode(ByVal SelString As String)
'
'    Dim strCD As String
'    Dim strNm As String
'
'    Select Case objMyList.Tag
'        Case "PtID"
'             txtPtId.Text = medGetP(SelString, 1, ";")
'             lblPtNm.Caption = medGetP(SelString, 2, ";")
'             Call EnableButton(True)
'        Case "DoctId"
'             txtDoctorId.Text = Trim(medGetP(SelString, 1, ";"))
'             lblDoctNm.Caption = Trim(medGetP(SelString, 2, ";"))
'        Case "DeptCd"
'             txtDeptCd.Text = Trim(medGetP(SelString, 1, ";"))
'        Case "WardId"
'             txtWardId.Text = Trim(medGetP(SelString, 1, ";"))
'
'    End Select
'
'End Sub

Private Sub objMyList_SelectedItem(ByVal pSelectedItem As String)
    Select Case objMyList.Tag
        Case TAG_WARD
            txtWardId.Text = objMyList.SelectedItems(0)
        Case TAG_DEPT
            txtDeptCd.Text = objMyList.SelectedItems(0)
        Case TAG_DOCT
            txtDoctorId.Text = objMyList.SelectedItems(0)
            lblDoctNm.Caption = objMyList.SelectedItems(1)
    End Select
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DELETE
            tblOrdSheet.Col = -1
            tblOrdSheet.Action = ActionDeleteRow
    End Select
End Sub

Private Sub optDiv_Click(Index As Integer)
    
    Call cmdClear_Click
    If Index = 0 Then
        tblCollect.Visible = True
        tblOrdSheet.Visible = False
        cmdByPass.Visible = False
    Else
        tblCollect.Visible = False
        tblOrdSheet.Visible = True
        cmdByPass.Visible = True
    End If
    
End Sub

Private Sub optOption_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtPtId.Text = "" Then
            txtPtId.SetFocus
            Exit Sub
        End If
        Call optOption_Click(Index)
        txtDoctorId.SetFocus
    End If
End Sub

Private Sub optOption_LostFocus(Index As Integer)
    If txtPtId.Text = "" Then
        txtPtId.SetFocus
        Exit Sub
    End If
End Sub

'% 처방명 또는 처방코드 입력
Private Sub tblOrdSheet_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim tmpIndex    As Integer
    Dim tmpStr      As String
    Dim Wdt         As Long
    Dim Hgt         As Long
    Dim X           As Long
    Dim Y           As Long
    Dim Ret         As Boolean

    If Col <> enORDSHEET.tcTESTNM Then Exit Sub

    With tblOrdSheet
        .Col = Col
        .Row = Row

        tmpIndex = medListFind(lstTestList, tblOrdSheet.Value)
        tmpStr = lstTestList.List(tmpIndex)


        ' 문자가 입력될때마다 유사어 찾기

        If tmpIndex = -1 Or UCase(tmpStr) <> UCase(.Value) Then
            Ret = .GetCellPos(Col, Row + 1, X, Y, Wdt, Hgt)
            If .Height - Y < lstTestList.Height Or Y < 0 Then
                Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                lstTestList.Top = .Top + Y - lstTestList.Height
            Else
                lstTestList.Top = .Top + Y
            End If
            If tmpIndex >= 0 Then
                medLockWindowUpdate (lstTestList.hwnd)

                lstTestList.ListIndex = tmpIndex
                medLockWindowUpdate (0&)
                If tmpIndex - lstTestList.TopIndex > 10 Then lstTestList.TopIndex = tmpIndex
            End If
            lstTestList.Visible = True
            lstTestList.ZOrder 0
        Else
            medLockWindowUpdate (lstTestList.hwnd)

            lstTestList.ListIndex = tmpIndex
            medLockWindowUpdate (0&)
            Call lstTestList_MouseDown(1, 0, 0, 0)
            lstTestList.Visible = False
        End If
    End With
End Sub

'% 처방항목 리스트가 떠 있고 아래화살표키를 눌렀을 경우 포커스 이동
Private Sub tblOrdSheet_KeyDown(KeyCode As Integer, Shift As Integer)

    With lstTestList
        If .Visible Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown:
                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
                    .SetFocus
                Case vbKeyUp, vbKeyPageUp:
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                    .SetFocus
                Case vbKeyEscape:
                    .Visible = False
                    'tblOrdSheet.SetFocus
            End Select
        End If
    End With

End Sub

'% 처방항목 리스트가 떠 있고 엔터키를 눌렀을 경우 항목 선택
Private Sub tblOrdSheet_KeyPress(KeyAscii As Integer)
    With tblOrdSheet
        If KeyAscii = vbKeyReturn And lstTestList.Visible Then
            DoEvents
            Call lstTestList_MouseDown(1, 0, 0, 0)
        End If
    End With
End Sub

'% 검체코드/희망채취일시 필드로 커서가 옮겨지면 검체리스트/날짜변경box 팝업
Private Sub tblOrdSheet_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim tmpTestCd As Variant
    Dim tmpSpcCd As Variant
    Dim tmpDate As Variant
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    If NewCol = enORDSHEET.tcTESTNM And lstTestList.Visible Then
        Cancel = True
        lstTestList.SetFocus
        Exit Sub
    End If
    If Col = enORDSHEET.tcTESTNM And lstTestList.ListIndex < 0 And lstTestList.Visible Then
        Cancel = True
        lstTestList.SetFocus
        Exit Sub
    End If

    If ActiveControl.Name = lstSpcList.Name Then Exit Sub

    If Col = enORDSHEET.tcSPCCD Then lstSpcList.Visible = False
    'If Col = enORDSHEET.tcREQDTTM Then fraDatePicker.Visible = False

    Select Case NewCol
    Case enORDSHEET.tcSPCCD:    ' 검체리스트
        If lstSpcList.Visible Then Exit Sub
        With tblOrdSheet
            .Row = NewRow: .Col = NewCol
            Ret = .GetText(enORDSHEET.tcTESTCD, NewRow, tmpTestCd)
            If tmpTestCd = "" Then Cancel = True: Exit Sub
            'Ret = .GetText(4, NewRow, tmpSpcCd)
            Ret = .GetCellPos(NewCol, NewRow + 1, X, Y, Wdt, Hgt)
            If Y > 0 Then
                lstSpcList.Top = .Top + Y
            Else
                Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
                lstSpcList.Top = .Top + Y - lstSpcList.Height
            End If
            Call objOrder.SpcList(tmpTestCd, lstSpcList)
            lstSpcList.Visible = True
            lstSpcList.ZOrder 0
            lstSpcList.SetFocus
            If lstSpcList.ListCount > 0 Then lstSpcList.ListIndex = 0
            DoEvents
        End With
    Case 7:   ' 희망채혈일시 입력
'        If fraDatePicker.Visible Then Exit Sub
'        With tblOrdSheet
'            .Row = NewRow: .Col = NewCol
'            Ret = .GetText(enORDSHEET.tcTESTCD, NewRow, tmpTestCd)
'            If tmpTestCd = "" Then Cancel = True: Exit Sub
'            Ret = .GetText(enORDSHEET.tcREQDTTM, NewRow, tmpDate)
'            Ret = .GetCellPos(NewCol, NewRow + 1, X, Y, Wdt, Hgt)
'            If Y > 0 Then
'                fraDatePicker.Top = .Top + Y
'            Else
'                Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
'                fraDatePicker.Top = .Top + Y - fraDatePicker.Height
'            End If
'            fraDatePicker.Visible = True
'            If tmpDate = "" Then
'                txtDatePicker.Value = GetSystemdate
'            Else
'                txtDatePicker.Value = tmpDate
'            End If
'            txtDatePicker.SetFocus
'            DoEvents
'        End With
    End Select

End Sub

'% 부서코드
Private Sub txtDeptCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(2)
End Sub

Private Sub txtDeptCd_LostFocus()
    If txtDeptCd.Text <> "" Then Call txtDeptCd_KeyPress(vbKeyReturn)
End Sub

Private Sub txtDoctorId_LostFocus()
    If txtDoctorId.Text <> "" Then Call txtDoctorId_KeyPress(vbKeyReturn)
End Sub


'% 병동ID
Private Sub txtWardId_GotFocus()
    With txtWardId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(3)
End Sub

'% SetFocus : 병동ID --> 병실ID
Private Sub txtWardId_KeyPress(KeyAscii As Integer)
'    Dim objData As clsBasisData
    Dim strData As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtWardId.Text = "" Then
            txtWardId.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            strData = GetWardNm(txtWardId.Text)
'            Set objData = Nothing
            
            If strData = "" Then
'            If Not objLisComCode.WardId.Exists(txtWardId.Text) Then
                MsgBox "병동 코드를 확인하세요.."
                txtWardId.Text = ""
                Call cmdHelpList_Click(3)
                Exit Sub
            End If
        End If
        txtRoomId.SetFocus
    End If

End Sub

'% 병실 ID
Private Sub txtRoomId_GotFocus()
    With txtRoomId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 병실ID --> 침상ID
Private Sub txtRoomId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtBedId.Enabled Then
        txtBedId.SetFocus
    End If
End Sub

'% 침상ID
Private Sub txtBedId_GotFocus()
    With txtBedId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 침상ID --> Order sheet
Private Sub txtBedId_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn AND dtpColDate.Enabled Then
'        dtpColDate.SetFocus
'    End If
End Sub


'% 부서코드
Private Sub txtDeptCd_GotFocus()
    With txtDeptCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 부서코드 --> Ward ID / ReceptNo
Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    Dim objData As clsBasisData
    Dim strData As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If ActiveControl.Name = txtPtId.Name Then Exit Sub
    If ActiveControl.Name = optOption(0).Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    If KeyAscii = vbKeyReturn Then
        If txtDeptCd.Text = "" Then
            txtDeptCd.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            strData = GetDeptNm(txtDeptCd.Text)
'            Set objData = Nothing
            
            If strData = "" Then
'            If Not objLisComCode.DeptCd.Exists(txtDeptCd.Text) Then
                MsgBox "부서 코드를 확인하세요.."
                txtDeptCd.Text = ""
                Call cmdHelpList_Click(2)
                Exit Sub
            End If
        End If
        If optOption(1).Value Then
            txtWardId.SetFocus
        Else
            If txtPtId.Text = "" Then
                txtPtId.SetFocus
                Exit Sub
            End If
            If txtDoctorId.Text = "" Then
                txtDoctorId.SetFocus
                Exit Sub
            End If
            If optOption(1).Value And (txtWardId.Text = "") Then
                txtWardId.SetFocus
                Exit Sub
            End If

        End If
    End If

End Sub

'% 의사ID
Private Sub txtDoctorId_GotFocus()
    With txtDoctorId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDoctorId_Change()
    lblDoctNm.Caption = ""
End Sub

'% Arrow Down --> 의사ID 리스트 팝업
Private Sub txtDoctorId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(1)
End Sub

'% SetFocus : 의사ID --> 부서코드
Private Sub txtDoctorId_KeyPress(KeyAscii As Integer)
'    Dim objData As clsBasisData
    
    
    If KeyAscii = vbKeyReturn Then
        If txtDoctorId.Text = "" Then
            lblDoctNm.Caption = ""
            txtDoctorId.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            lblDoctNm.Caption = GetEmpNm(txtDoctorId.Text) 'GetEmpName(txtDoctorId.Text)
'            Set objData = Nothing
            
            If lblDoctNm.Caption = "" Then
                MsgBox "처방의 코드를 확인하세요.."
                txtDoctorId.Text = ""
                Call cmdHelpList_Click(1)
                Exit Sub
            End If
        End If
        txtDeptCd.SetFocus
    End If
End Sub

'% 환자ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPtId_Change()
    If Not blnClearFg Then Call ClearRtn(False)
End Sub

'% 환자ID를 Key로 데이타 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then optOption(0).SetFocus

End Sub

'% 환자ID를 Key로 데이타 검색
Private Sub txtPtId_LostFocus()

    Dim blnRst As Boolean

    Call ClearData
    
    
    
    '결과등록을 위한 초기화
    Set objDic = New clsDictionary
    objDic.Clear
    objDic.FieldInialize "seq", "workarea,accdt,accseq,donefg"
    fraAcc.Enabled = False
    cmdAcc.Enabled = False
    cmdNextData(0).Enabled = False
    cmdNextData(1).Enabled = False
    cmdSave.Enabled = False
    
    
    If txtPtId.Text = "" Then Exit Sub


    txtPtId.Text = UCase(txtPtId.Text)
    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)


    blnRst = objPatient.GETPatient((txtPtId.Text))

    If Not blnRst Then
        MsgBox "등록되지 않은 환자입니다. ID를 확인하세요 ! ", vbExclamation + vbOKOnly, "처방등록"
        txtPtId.Text = ""
        DoEvents
        txtPtId.SetFocus
        Exit Sub
    End If

    Call DisplayPtInfo
    If optDiv(0).Value Then DisplayOrder
    Call EnableButton(True)
    
    '감염관리 표시
    Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)

    blnClearFg = False
    optOption(0).SetFocus

End Sub
'% 검색한 처방을 테이블에 디스플레이 한다.

Public Function SqlReadWardOrder(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                  Optional ByVal Bussdiv As String = "") As String
 
    Dim tmpStr  As String
    Dim tmpStr1 As String
    Dim strSQL  As String

    tmpStr = "": tmpStr1 = ""
    
    If Bussdiv <> "" Then tmpStr = tmpStr & " AND  " & DBW("a.bussdiv = ", Bussdiv)     '외래병동구분
    
    If Bussdiv = "1" Then '외래
        tmpStr = tmpStr & " AND  " & DBW("a.orddt = ", ReqDt)
    Else
        tmpStr1 = " AND a.reqdt||a.reqtm<='" & ReqDt & ReqTm & "'"
    End If

    strSQL = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
            "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
            "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                     FUNC_CONVERT("int", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
            "        '' as bedid, '' as fzfg " & _
            " FROM " & T_LAB001 & " c, " & T_LAB032 & " d, " & T_HIS005 & " e, " & _
                       T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
            " WHERE " & DBW("a.ptid = ", PtId) & _
            " AND    a.donefg = '0' " & _
            " AND  " & DBW("a.orddiv = ", LIS_ORDDIV) & tmpStr1 & _
            " AND    b.ptid  = a.ptid " & _
            " AND    b.orddt = a.orddt " & _
            " AND    b.ordno = a.ordno " & _
            " AND  " & DBW("b.ordcd=", P_POCCode) & _
            " AND    b.donefg = '0' " & tmpStr & _
            " AND   (b.dcfg = '' or b.dcfg is null) " & _
            " AND    c.testcd  = b.ordcd " & _
            " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
            " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
            " AND  " & DBJ("d.cdval1 =* b.spccd") & _
            " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
            " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
            " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) "

 

    SqlReadWardOrder = strSQL

    SqlReadWardOrder = SqlReadWardOrder & " ORDER BY ColTm, orddt, ordno, ordcd "                          '<< D/C 처방 제외 >>
      
     
End Function
Private Sub DisplayOrder()
   
    Dim i           As Integer
    Dim SqlStmt     As String
    Dim tmpRs       As Recordset
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim strErChk    As String
    Dim strOrdDiv   As String
    
    Dim objProInSts As clsProgress

On Error GoTo NoData
   
    Call medClearTable(tblCollect)
    tblCollect.Enabled = False
    Set objProInSts = New clsProgress
    With objProInSts
        .Container = Me
        .Left = LisLabel1.Left + 5
        .Top = LisLabel1.Top + 5
        .Width = tblCollect.Width - 10
        .Height = LisLabel1.Height - 10
        .Message = "해당환자의 처방 내역을 검색 중입니다...."
        .Max = 90
        .Value = 10
        
'        .SetMyForm Me
'        .Choice = True
'        .ForeColor = &HFA8B10
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = LisLabel1.Height - 10
'        .MSG = "해당환자의 처방 내역을 검색 중입니다...."
'        .Max = 90
'        .Min = 0
'        .Value = 10
        DoEvents
    End With

    DoEvents
'    txtMesg.Text = ""
    
    ' 처방내역 검색

    tmpDate = Format(GetSystemDate, CS_DateDbFormat)
    tmpTime = "235959"
 
    SqlStmt = SqlReadWardOrder(txtPtId.Text, tmpDate, tmpTime, enBussDiv.BussDiv_OutPatient)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then
        Set tmpRs = Nothing
        Set objProInSts = Nothing
       
        MsgBox objPatient.ptnm & " 님의 처방내역이 없습니다", vbInformation, "간호사 채취"
        
        Exit Sub
    End If
    
    With tblCollect
       
        .ReDraw = False
        objProInSts.Max = tmpRs.RecordCount
        .Row = -1
        .Col = 2: .COL2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
             
        For i = 1 To tmpRs.RecordCount

            objProInSts.Value = i
            If .DataRowCnt + 1 > .MaxRows Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
            
            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateShortMask)    '처방일
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)     '처방의
                SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    '처방번호
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '검체
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) '처방의
            End If
            
            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    '처방의
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    '처방번호
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '검체
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) '처방의
            End If
            
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '검체
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            End If
            
            If SvOrdDoct <> Trim("" & tmpRs.Fields("DoctNm").Value) Then
                .Col = enCOLLIST.tcDOCTNM: .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    '처방의
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)
            End If

            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")   '건물별 응급가능 여부
            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")   '건물별 검사가능 여부


            '***건물정보 사용
            If P_ApplyBuildingInfo Then
               If Trim(tmpRs.Fields("StatFg").Value) = "1" Then
                   '**응급검사 가능
                   If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                       '** 중앙/안이센터에서 응급검사가 발생하면.. --> 응급센터로...
                       If ObjSysInfo.BuildingCd = CentralLab Or _
                          ObjSysInfo.BuildingCd = AneLab Then
                           .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                           .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
                       '** 해당건물에서 응급검사 가능함
                       Else
                           .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                           .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
                       End If
                       .Col = enCOLLIST.tcSTATFLAG: .Text = "1"       'StatFg
                       GoTo DataSet
                   Else
                   '*******************************************************************************************************
                   '** 여성/심장센터 : 응급검사가 가능하지 않을경우 응급실에서 검사가 가능하면 응급실로, 아니면 중앙으로...
                   '*******************************************************************************************************
                       '** 여성/심장센터에서 응급검사가 발생하면..
                       If ObjSysInfo.BuildingCd = WomLab Or ObjSysInfo.BuildingCd = HrtLab Then
                           '** 응급실에서 응급검사 가능 --> 응급센터로...
                           If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then
                               .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                               .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
                               .Col = enCOLLIST.tcSTATFLAG:   .Text = "1"   'StatFg
                               GoTo DataSet
                           End If
                       End If
                   '*******************************************************************************************************
                   End If
               End If
    
               .Col = enCOLLIST.tcSTATFLAG: .Text = "0"          'StatFg

               '**일반검사가능
               If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                   
                   .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                   
                   .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm

               '**일반검사 불가능 --> 중앙검사실로...
               Else
                   .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
                   
                   .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
               End If
    
            '***건물정보 사용하지 않음
                Else
                    .Col = enCOLLIST.tcBUILDCD:  .Text = ObjSysInfo.BuildingCd
                    .Col = enCOLLIST.tcBUILDNM:  .Text = ObjSysInfo.BuildingNm
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(tmpRs.Fields("StatFg").Value)
                End If
          
DataSet:
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & tmpRs.Fields("TestNm").Value)     '처방명
                                        .ForeColor = &H553755
            .Col = enCOLLIST.tcSTATFG:  .Text = IIf("" & tmpRs.Fields("StatFg").Value = "0", "", "Y") '응급여부
                                        .ForeColor = DCM_LightRed                                '빨간색
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)      '희망채취일시
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & tmpRs.Fields("OrdDt").Value)      '처방일
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & tmpRs.Fields("OrdNo").Value)      '처방번호
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & tmpRs.Fields("OrdSeq").Value)     '처방Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & tmpRs.Fields("OrdCd").Value)      '검사코드

            Dim strLabDiv As String
            
            strLabDiv = GetLabDiv(.Text)
'            Call objLisComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = strLabDiv 'objLisComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & tmpRs.Fields("SpcCd").Value)      '검체코드

'            Call objLisComCode.LisSpc.KeyChange(.Text)
            Dim strLabRng As String
            Dim strSpcAbbr As String
            
            Call GetSpcInfo(.Text, strSpcAbbr, strLabRng)
            
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & tmpRs.Fields("spcnm5").Value)         '검체약어명
            .Col = enCOLLIST.tcLABRANGE: .Text = strLabRng 'objLisComCode.LisSpc.Fields("labrange")     '미생물접수번호범위

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & tmpRs.Fields("WorkArea").Value)  'WorkArea
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & tmpRs.Fields("StoreCd").Value)   '보관코드
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & tmpRs.Fields("TestDiv").Value)   '검사구분
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & tmpRs.Fields("MultiFg").Value)   '복수검체여부
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & tmpRs.Fields("SpcGrp").Value)    '검체군
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & tmpRs.Fields("OrdDoct").Value)   '처방의
                                         '처방의명
                                         txtDoctorId.Text = .Text
                                         If .Text <> "" And lblDoctNm.Caption = "" Then
                                            lblDoctNm.Caption = Trim("" & tmpRs.Fields("DoctNm").Value)
                                         End If
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & tmpRs.Fields("MajDoct").Value)   '주치의
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & tmpRs.Fields("DeptCd").Value)    '진료과
                                        txtDeptCd.Text = .Text
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & tmpRs.Fields("AbbrNm5").Value)    '약어명
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & tmpRs.Fields("LabelCnt").Value)   '라벨출력장수
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & tmpRs.Fields("ReceptNo").Value)   '영수증번호
                                        .ForeColor = vbRed

            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & tmpRs.Fields("WardId").Value)     '병동
                                        txtWardId.Text = GetWardNm(.Text) 'objLisComCode.WardId.Fields("wardid")
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & tmpRs.Fields("hosilid").Value)     '병실
                                        txtRoomId.Text = .Text
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & tmpRs.Fields("roomid").Value)      '병상
                                        txtBedId.Text = .Text
            
            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & tmpRs.Fields("FzFg").Value)       '동결절편
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & tmpRs.Fields("OrdDiv").Value)     '처방구분

'            '진료부서 Remark
'            If Trim("" & tmpRs.Fields("Mesg").Value) <> "" Then
'                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & tmpRs.Fields("OrdNo").Value), "##") & " - "
'                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("TestNm").Value) & vbCrLf
'                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("Mesg").Value) & vbCrLf
'            End If

            tmpRs.MoveNext
        Next
'        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
       
    End With
    
    Set objProInSts = Nothing
    
NoData:
    Set tmpRs = Nothing
   
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    GetLabDiv = RS.Fields("field2").Value & ""
    
    Set RS = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    vSpcAbbr = RS.Fields("spcabbr").Value & ""
    vLabRng = RS.Fields("labrange").Value & ""
    
    Set RS = Nothing
End Sub

'Private Function GetWardNm(ByVal vWardId As String)
'    Dim objData As clsBasisData
'
'    Set objData = New clsBasisData
'    GetWardNm = objData.GetWardNm(vWardId)
'    Set objData = Nothing
'
'End Function
'% 환자정보 클래스 objPatient 로부터 기본정보를 Screen에 Display한다.
Private Sub DisplayPtInfo()

    With objPatient
        
        lblPtNm.Caption = .ptnm
        lblAgeDiv.Caption = .AGEDIV
        lblSex.Caption = .SEXAGE
    
        txtDeptCd.Text = .DeptCd ' .DeptCd
        txtDoctorId.Text = .MajDoct
        lblDoctNm.Caption = .DoctNm
        txtWardId.Text = .WardId
        txtRoomId.Text = .ROOMID
        txtBedId.Text = .BedID
        
        If .INADMISSION Then
            optOption(1).Value = True
        Else
            optOption(0).Value = True
        End If
    
        lblSex.Caption = .SEXNM
        lblAge.Caption = .Age
        lblDob.Caption = Format(.Dob, CS_DateMask)
    End With

End Sub

'% Order sheet에서 Row 삭제
'Private Sub mnuDelete_Click()
'    tblOrdSheet.Col = -1
'    tblOrdSheet.Action = ActionDeleteRow
'End Sub

'% 처방유형에 따라 희망채혈시간 컨트롤 변경
Private Sub optOption_Click(Index As Integer)

    objOrder.OrdDiv = LIS_ORDDIV   'Choose(Index + 1, "S", "W", "L")
    If Index = 2 Then
    
    Else
        If Index = 0 Then
            cmdHelpList(3).Enabled = False
            txtWardId.Enabled = False
            txtRoomId.Enabled = False
            txtBedId.Enabled = False
        ElseIf Index = 1 Then
            cmdHelpList(3).Enabled = True
            txtWardId.Enabled = True
            txtRoomId.Enabled = True
            txtBedId.Enabled = True
        End If
    End If

End Sub

'% 오른쪽 버튼 클릭 --> Delete Row
Private Sub tblOrdSheet_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    Dim lngOldColor As Long

    tblOrdSheet.OperationMode = OperationModeRead
    tblOrdSheet.Row = Row
    tblOrdSheet.Col = -1
    lngOldColor = tblOrdSheet.BackColor
    tblOrdSheet.BackColor = DCM_LightGray

'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    frmControls.mnuSub1.Visible = False
'    frmControls.mnuSub2.Visible = False
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup

    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_DELETE, "Delete"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing
    
    tblOrdSheet.Row = Row
    tblOrdSheet.Col = -1
    tblOrdSheet.BackColor = lngOldColor
    tblOrdSheet.OperationMode = OperationModeNormal
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing

End Sub

'% 종료
Private Sub cmdExit_Click()
    Unload Me
    Set objSqlStmt = Nothing
    Set objPatient = Nothing
    Set objOrder = Nothing
    Set objCollect = Nothing
    Set objAccess = Nothing
    Set frm101Order = Nothing
End Sub

Private Function ValidationCheck() As Boolean

    Dim CheckOrder As Long

    ValidationCheck = True

    If tblOrdSheet.DataRowCnt = 0 Then GoTo Err_Trap

    If txtDoctorId.Text = "" Then
        MsgBox "처방의ID를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtDoctorId.SetFocus
        GoTo Err_Trap
    End If
    If txtDeptCd.Text = "" Then
        MsgBox "진료과를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtDeptCd.SetFocus
        GoTo Err_Trap
    End If
    If (txtWardId.Text = "") And txtWardId.Enabled Then
        MsgBox "병동ID를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtWardId.SetFocus
        GoTo Err_Trap
    End If

    CheckOrder = objOrder.CheckSameOrder(tblOrdSheet)
    If CheckOrder > 0 Then
        tblOrdSheet.Row = CheckOrder
        tblOrdSheet.Col = enORDSHEET.tcTESTNM
        tblOrdSheet.Action = ActionActiveCell
        MsgBox "중복처방입니다. 희망채혈시간을 조정하세요..", vbExclamation, "입력오류"
        tblOrdSheet.SetFocus
        GoTo Err_Trap
    End If

    Exit Function

Err_Trap:
    ValidationCheck = False

End Function

'% 클래스 objOrder의 데이타 속성을 채우고 SaveData 메쏘드를 Call하여
'% 데이타베이스로 저장한다.
'% 채혈/접수 클래스를 통하여 채혈/접수 내역을 생성하고 Barcode를 출력한다.
Private Function SaveOrder(Optional ByRef objPrgBar As clsProgress = Nothing) As Boolean

    Dim i As Integer
    Dim StartOrdNo As Integer
    Dim OrderSuccess As Boolean

    With objOrder

        .PtId = txtPtId.Text
        .ordDt = Format(GetSystemDate, CS_DateDbFormat)
        .OrdTm = Format(GetSystemDate, CS_TimeDbFormat)
        
        If optOption(1).Value Then
            .Bussdiv = enBussDiv.BussDiv_InPatient    '병동
            '## 병동ID(HIS002)는 임시로 현재일 Setting ...
            .BedInDt = objPatient.BedInDt
            .DeptCd = txtDeptCd.Text
            .MajDoct = objPatient.MajDoct
            .WardId = txtWardId.Text
            .HosilId = txtRoomId.Text
            .ROOMID = txtBedId.Text
        Else
            .Bussdiv = enBussDiv.BussDiv_OutPatient   '외래
            .BedInDt = ""
            .DeptCd = txtDeptCd.Text
            .MajDoct = txtDoctorId.Text
            .WardId = ""
            .HosilId = ""
            .ROOMID = ""
        End If
        .OrdDoct = txtDoctorId.Text
        .ReceptNo = ""
        .EntId = ObjSysInfo.EmpId
        .EntDt = Format(GetSystemDate, CS_DateDbFormat)
        .EntTm = Format(GetSystemDate, CS_TimeDbFormat)
        .DoneFg = enStsCd.StsCd_LIS_Order
        .RepeatFg = ""
        .OrgAccNo = ""
        .SpOrdDiv = ""
        .OrdDiv = LIS_ORDDIV

        Call .MoveData(tblOrdSheet)                     '클래스로 데이타 Move
        OrderSuccess = .SaveData(StartOrdNo, objPrgBar) 'Database로 저장

    End With

    If Not OrderSuccess Then
        SaveOrder = False
        Exit Function
    End If

    '처방번호 Display
    With tblOrdSheet
        .Col = 1
        For i = 1 To .DataRowCnt
            .Row = i
            .Value = Val(.Value) + StartOrdNo
        Next
    End With
    SaveOrder = True

End Function

'% 발생한 처방데이타를 기준으로 채혈접수내역을 생성하기 위해
'% 모든 데이타를 Array로 Assign한다.
Private Sub ReadyToCollect()

    
    Dim tmpData()   As String
    Dim strDOB      As String
    Dim i           As Integer
    
    
    With objCollect
        ReDim tmpData(0 To 16)
        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)      '검체년도
        tmpData(1) = objPatient.PtId                                '환자ID
        tmpData(2) = objPatient.ptnm
        tmpData(3) = objPatient.Sex                                 '성별
        If IsDate(lblDob.Caption) Then                              '환자일령
            tmpData(4) = DateDiff("y", lblDob.Caption, GetSystemDate)
        Else
            tmpData(4) = DateDiff("y", Mid(lblDob.Caption, 1, 4) & "-01-01", GetSystemDate)
        End If
        tmpData(5) = objPatient.BedInDt                             '입원일
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)     '입력일
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)     '입력시간
        tmpData(8) = ObjSysInfo.EmpId                               '입력자
        tmpData(9) = ""                                             '원접수번호
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)    '채혈일
        .ColTm = Format(GetSystemDate, "hhmmss")
        tmpData(11) = ObjSysInfo.EmpId                              '채혈자
        tmpData(12) = txtWardId.Text                                '병동ID
        tmpData(13) = txtRoomId.Text                                '병실ID
        tmpData(14) = ""                                            '병실ID
        tmpData(15) = txtBedId.Text                                 '침상ID
        tmpData(16) = ObjSysInfo.BuildingCd                         '** 채혈이 수행되는 건물코드
        Call .SetColData(tmpData)
    End With

    If optDiv(1).Value Then
        With tblOrdSheet
            ReDim tmpData(0 To .MaxCols)
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = enORDSHEET.tcBUILDCD:    tmpData(0) = .Value     'Delivery Location
                .Col = enORDSHEET.tcWORKAREA:   tmpData(1) = .Value     'WorkArea
                .Col = enORDSHEET.tcSPCCD:      tmpData(2) = .Value     'SpcCd
                .Col = enORDSHEET.tcSTORECD:    tmpData(3) = .Value     'StoreCd
                .Col = enORDSHEET.tcSTATCHK:    tmpData(4) = .Value     'StatFg
                .Col = enORDSHEET.tcREQDTTM:    tmpData(5) = .Value     'ReqColDate
                .Col = enORDSHEET.tcTESTDIV:    tmpData(6) = .Value     'TestDiv
                .Col = enORDSHEET.tcMULTIFG:    tmpData(7) = .Value     'MultiFg
                .Col = enORDSHEET.tcSPCGRP:     tmpData(8) = .Value     'SpcGrp
                tmpData(9) = Format(GetSystemDate, CS_DateDbFormat) '처방일을 희망채혈일로.. 2000/04/03 by 정미경
                .Col = enORDSHEET.tcORDNO:      tmpData(10) = .Value    'OrdNo
                .Col = enORDSHEET.tcORDSEQ:     tmpData(11) = .Value    'OrdSeq
                .Col = enORDSHEET.tcTESTCD:     tmpData(12) = .Value    'OrdCd
                tmpData(13) = txtDeptCd.Text                            '진료과
                tmpData(14) = txtDoctorId.Text                          '처방의
                tmpData(15) = objPatient.MajDoct                        '주치의
                .Col = enORDSHEET.tcABBRNM:     tmpData(16) = .Value    '약어명
                .Col = enORDSHEET.tcBARCNT:     tmpData(17) = .Value    '라벨출력장수
                .Col = enORDSHEET.tcLABDIV:     tmpData(18) = .Value    '접수번호부여기준
                .Col = enORDSHEET.tcSPCABBR:    tmpData(19) = .Value    '검체약어명
                .Col = enORDSHEET.tcLABRANGE:   tmpData(20) = .Value    '미생물접수번호범위
    
                Call objCollect.AddLabCollect(tmpData)
            Next
        End With
    Else
        With tblCollect
        
            ReDim tmpData(0 To 20)
'            .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
'                            .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
'                            .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
                .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
                .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
                .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
                .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
                .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .Value        'ReqColDate
                .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value        'TestDiv
                .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value        'MultiFg
                .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value        'SpcGrp
                .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value        'OrdDt
                .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value       'OrdNo
                .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value       'OrdSeq
                .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value       'OrdCd
                .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .Value       '진료과
                .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value       '처방의
                .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value       '주치의
                .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value       '검사 약어명
                .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value       '라벨출력장수
                .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value       'LabDiv
                .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value       '검체약어명
                .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value       '미생물접수번호범위
                Call objCollect.AddLabCollect(tmpData)
            Next
        End With
    End If
End Sub


'% Clear Routine : 환자정보 클래스, 처방 클래스 및 각 콘트롤들 초기화
Sub ClearRtn(Optional ByVal blnAll As Boolean = True)
    
    If blnAll Then txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDob.Caption = ""
    optOption(0).Value = True
    txtDoctorId.Text = ""
    lblDoctNm.Caption = ""
    txtDeptCd.Text = ""
    txtWardId.Text = ""
    txtRoomId.Text = ""
    txtBedId.Text = ""
    cmdByPass.Enabled = False
    fraPass.Visible = False

    Call EnableButton(False)
    tblCollect.Visible = True
    tblOrdSheet.Visible = False
    cmdByPass.Visible = False
    Call medClearTable(tblCollect)
'    Call optDiv_Click(0)
    With tblOrdSheet
        .MaxRows = 0
        .MaxRows = 50
        .Row = -1
        .Col = enORDSHEET.tcORDNO: .COL2 = enORDSHEET.tcORDNO
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
    End With
    
    Set objPatient = Nothing
    Set objPatient = New clsPatient

    Set objCollect = Nothing
    Set objCollect = New clsLISCollectioin
    
    Set objOrder = Nothing
    Set objOrder = New clsLISOrder
    objOrder.BuildingCd = ObjSysInfo.BuildingCd
    objOrder.BuildingNm = ObjSysInfo.BuildingNm
    objOrder.BuildingNo = ObjSysInfo.BuildingNo
    
    Set objAccess = Nothing
    Set objAccess = New clsLISAccession
    '결과등록관련
    mskAccNo.BackColor = DCM_LightGray
    mskAccNo.Enabled = False

End Sub


Private Sub EnableButton(ByVal ValFg As Boolean)
    tblOrdSheet.Enabled = ValFg
    cmdByPass.Enabled = ValFg
End Sub

Private Sub txtWardId_LostFocus()
    If txtWardId.Text <> "" Then Call txtWardId_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdAccession_Click()
    Dim i       As Integer
    Dim Success As Boolean
    Dim objProgress As clsProgress

    MouseRunning  '13


    If optDiv(0).Value Then
        If tblCollect.DataRowCnt = 0 Then Exit Sub
        Set objProgress = New clsProgress
        objProgress.Container = MainFrm.stsBar
'        objProgress.CaptionOn = False
'        objProgress.Visible = True
'        objProgress.Min = 0
        objProgress.Max = tblCollect.DataRowCnt * (2 + 1)
    Else
        If tblOrdSheet.DataRowCnt = 0 Then Exit Sub
        If Not ValidationCheck Then Exit Sub
        Set objProgress = New clsProgress
'        objProgress.CaptionOn = False
'        objProgress.Visible = True
'        objProgress.Min = 0
        objProgress.Message = "환자의 처방내역을 저장하고 있습니다."
        objProgress.Max = tblOrdSheet.DataRowCnt * (2 + 1)
        With tblOrdSheet
            .SortBy = SortByRow
            .SortKey(1) = enORDSHEET.tcBUILDCD      'DeliveryLocation
            '.SortKey(2) = 12  '검사구분  --> 제외 1999.10.08 by 김미경
            .SortKey(2) = enORDSHEET.tcREQDTTM      '희망채취시간
            .SortKey(3) = enORDSHEET.tcWORKAREA     'WorkArea
            .SortKey(4) = enORDSHEET.tcSPCCD        '검체코드
            .SortKey(5) = enORDSHEET.tcSTORECD      '보관구분
            .SortKey(6) = enORDSHEET.tcSTATCHK      '응급여부
            .SortKey(7) = enORDSHEET.tcTESTCD       '검사코드
            .SortKeyOrder(1) = SortKeyOrderAscending
            .SortKeyOrder(2) = SortKeyOrderAscending
            .SortKeyOrder(3) = SortKeyOrderAscending
            .SortKeyOrder(4) = SortKeyOrderAscending
            .SortKeyOrder(5) = SortKeyOrderAscending
            .SortKeyOrder(6) = SortKeyOrderAscending
            .SortKeyOrder(7) = SortKeyOrderAscending
            .Col = 1:   .COL2 = .MaxCols
            .Row = 0:  .Row2 = .MaxRows
            .Action = ActionSort
        End With
        'Button 1  : 처방
        Success = SaveOrder(objProgress)                    '처방내역 저장
    
        If Success = False Then
            Set objProgress = Nothing
            MsgBox "처방내역 저장중 오류가 발생했습니다. ", vbCritical, "오류발생"
            GoTo Exit_Routine
        End If
    End If
    
    objProgress.Value = 0

    DoEvents

    'Button 2  : 채혈수행

    objProgress.Message = "채혈 Procedure를 수행하고 있습니다."

    Call ReadyToCollect                             '채혈준비
    Success = objCollect.DoCollection(objProgress)    '채혈수행

    If Success = False Then
        Set objProgress = Nothing
        MsgBox "채혈내역 생성중 오류가 발생했습니다. ", vbCritical, "오류발생"
        GoTo Exit_Routine
    End If

    'Button 3:  접수수행
    objProgress.Message = "접수 Procedure를 수행하고 있습니다."
    objDic.Sort = False
    With objCollect
        If .CollectDone Then
            Dim pWorkArea   As String
            Dim pAccDt      As String
            Dim pAccSeq     As Long
            Dim kk          As Integer
            
            For i = 1 To .ColCount
                objProgress.Message = "접수 Procedure를 수행하고 있습니다. (" & CStr(i) & "/" & CStr(.ColCount) & ")"
                Call .GetLabNumbers(i, pWorkArea, pAccDt, pAccSeq)
                Success = objAccess.DoAccession(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
                If Success Then
                    kk = kk + 1
                    objDic.AddNew kk, pWorkArea & COL_DIV & pAccDt & COL_DIV & pAccSeq & COL_DIV & "0"
                End If
                If Not Success Then Exit For
                If objProgress.Value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
                objProgress.Value = objProgress.Value + 1
                DoEvents
            Next
        End If
    End With
    objDic.Sort = True
    
    If Success Then
        objProgress.Value = objProgress.Max
        Set objProgress = Nothing
        MsgBox "정상적으로 접수되었습니다.", vbInformation + vbOKOnly, "Info"
        '결과등록
        If objDic.RecordCount > 0 Then
            
            objDic.MoveFirst
            Call LabNoResult(objDic.Fields("seq"), objDic.Fields("workarea"), _
                             objDic.Fields("accdt"), objDic.Fields("accseq"))
        End If
    Else
        Set objProgress = Nothing
        MsgBox "접수처리중 오류가 발생했습니다.", vbCritical + vbOKOnly, "Info"
    End If
    '
Exit_Routine:
    MouseDefault
    Set objProgress = Nothing
    
    If optDiv(1).Value Then
        With tblOrdSheet
            .MaxRows = 0
            .MaxRows = 50
            .Row = -1
            .Col = enORDSHEET.tcORDNO: .COL2 = enORDSHEET.tcORDNO
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        End With
    Else
        Call medClearTable(tblCollect)
    End If
    'Call cmdClear_Click
End Sub


'==============================='결과등록 관련=================================
Private Sub lvwCount_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sSEQ As String
    Dim sWorkArea As String
    Dim sAccDt As String
    Dim sAccSeq As String
    
    sSEQ = Item.Text
    sWorkArea = medGetP(Item.SubItems(1), 1, "-")
    sAccDt = medGetP(Item.SubItems(1), 2, "-")
    sAccSeq = medGetP(Item.SubItems(1), 3, "-")
    
    If Item.SubItems(2) = "미확" Then
        Call LabNoResult(sSEQ, sWorkArea, sAccDt, sAccSeq)
    Else
        MsgBox "결과확인된 접수번호입니다.", vbInformation + vbOKOnly, "Info"
    End If
    fraAcc.Visible = False
End Sub
Private Sub LabNoResult(ByVal SEQ As String, ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String)
    
    Call ClearData
    ssRst.MaxRows = 0
    cmdNextData(0).Enabled = True
    cmdNextData(1).Enabled = True
    cmdAcc.Enabled = True
    cmdSave.Enabled = True
    cmdCloseFra.Enabled = True
    fraAcc.Enabled = True

    mskAccNo.Mask = String(Len(WorkArea), "&") & "-"
    mskAccNo.Mask = mskAccNo.Mask & String(Len(Mid(AccDt, 3)), "#") & "-"
    mskAccNo.Mask = mskAccNo.Mask & String(Len(AccSeq), "#")
    
    mskAccNo.Mask = "&&-######-#"

    mskAccNo.Text = WorkArea & "-" & Mid(AccDt, 3) & "-" & AccSeq
    mskAccNo.Tag = SEQ
    Call Data_Load
    
End Sub

Private Sub EditData()
    ssRst.Enabled = True
    mskAccNo.BackColor = DCM_LightGray
    cmdSave.Enabled = True
    fraComment.Enabled = True
    lblCapRemark.Enabled = True
    rtfComment.BackColor = &HF1F5F4     'vbWhite
End Sub
Private Sub CmdTemplete(ByVal blnVisible As Boolean)
    cmdTextTemplete.Enabled = blnVisible
    cmdRemarkTemplete.Enabled = blnVisible
    cmdCommentTemplete.Enabled = blnVisible
End Sub

'결과입력다음 버튼
Private Sub cmdNextData_Click(Index As Integer)
    Dim sKey As String
    
    If Index = 0 Then
        sKey = Val(mskAccNo.Tag) - 1
    Else
        sKey = Val(mskAccNo.Tag) + 1
    End If
    

    If objDic.Exists(sKey) Then
        objDic.KeyChange sKey
        mskAccNo.Mask = String(Len(objDic.Fields("workarea")), "&") & "-"
        mskAccNo.Mask = mskAccNo.Mask & String(Len(Mid(objDic.Fields("accdt"), 3)), "#") & "-"
        mskAccNo.Mask = mskAccNo.Mask & String(Len(objDic.Fields("accseq")), "#")
        
        mskAccNo.Mask = "&&-######-#"
    
        mskAccNo.Text = objDic.Fields("workarea") & "-" & _
                        Mid(objDic.Fields("accdt"), 3) & "-" & _
                        objDic.Fields("accseq")
        
        If objDic.Fields("donefg") = "0" Then
            mskAccNo.Tag = objDic.Fields("seq")
            Call Data_Load
        End If
    Else
        MsgBox "Bypass화면을 통하여 접수된 검체가 없습니다.", vbInformation + vbOKOnly, "Info"
        Call ClearData
    End If

End Sub

'결과입력대상 Load
Public Sub Data_Load()
    Dim strBk As String
'
    strBk = mskAccNo.Text
    '
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
   '
    PtResultLoad Trim(mskAccNo.FormattedText)
    If objPtInfo.TestCount > 0 Then
        ClearFg = False
        Call EditData
        lblErr.Caption = ""
        SendKeys "{TAB}"
        ssRst.Row = 1
        ssRst.Col = objPtInfo.SSCol("RESULT")
        ssRst.Action = ActionActiveCell
    Else
        mskAccNo.Text = strBk
        ssRst.Visible = True
        MsgBox "해당 접수번호는 입력할 검사가 없습니다.", vbCritical + vbOKOnly, "Info"
        Call ClearData
    End If
End Sub
Private Sub PtResultLoad(ByVal strAccNo As String)
    ssRst.Visible = False
    DoEvents
    
    MouseRunning
    
    Set objPtInfo.prgBar = prgRst
    objPtInfo.PrgBarInit
    With objPtInfo
        .PtType = RESULT_BY_ACCESSION             '/* 결과등록 유형, 반드시 셋팅 해야 됨./
        .AccNo = strAccNo       '/* 접수번호, 반드시 셋팅 해야 됨./
        
        .LoadTable , ObjMyUser.EmpId
        
        If .TestCount > 0 Then
            CmdTemplete True
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
            .GetResultSpread ssRst, RESULT_BY_ACCESSION
        End If
    End With
    
    Dim ii As Integer
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 5: .ForeColor = DCM_LightRed: .FontBold = True
        Next
        .Row = 1
        .Col = 2
        .Action = ActionActiveCell
    End With
    
    
    MouseDefault
    
    objPtInfo.PrgBarClear
    DoEvents
End Sub
Public Sub ClearData()
    ClearFg = True
    mskAccNo.Mask = "&&-######-####"
    mskAccNo.Text = "__-______-____"
    lblErr.Caption = ""

    ssRst.MaxRows = 0
    ssRst.Enabled = False
    mskAccNo.BackColor = vbWhite
    cmdSave.Enabled = False
    CmdTemplete False
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
    fraComment.Enabled = False
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub
Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub
Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
   
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    With clsTemplete
        .Show
        If pintMode = 0 Then
            .lblName = "Edit " & strTitle
        Else
            .lblName = "Modify " & strTitle
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
Private Sub cmdRemarkTemplete_Click()
    
    Dim SqlStmt As String

    Set objCodeList = Nothing
'    Set objCodeList = New clspopuplist
    Set objCodeList = New clsPopUpList

    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    With objCodeList
        .FormCaption = "Remark"
        .HideColumnHeaders = True
        .HideSearchTool = True
        .LoadPopUp SqlStmt
    
'        Set .MyDb = dbconn
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = Me.cmdRemarkTemplete.Top + 5600
'        .Left = Me.cmdRemarkTemplete.Left + 200
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< 없 음 > ", 2, 1
    End With

End Sub
'Comment처리
Private Sub clsTemplete_CopyTemplete()
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

'결과 리마크 표시
'Private Sub objCodeList_SendCode(ByVal SelString As String)
'    If Not IsNull(SelString) And SelString <> "" Then
'        Select Case objCodeList.Tag
'            Case "Remark":
'                objPtInfo.RmkCd = medGetP(SelString, 1, vbTab)
'                If Trim(objPtInfo.RmkCd) <> "" Then
'                    objPtInfo.RmkNm = medGetP(SelString, 2, vbTab)
'                Else
'                    objPtInfo.RmkNm = ""
'                End If
'                rtfRemark.Text = objPtInfo.RmkNm
'        End Select
'    End If
'    Set objCodeList = Nothing
'End Sub


Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = 18
    ssRst.Value = ""
    If ClearFg Then Exit Sub
'    gblnModify = True
End Sub

Private Sub ssRst_GotFocus()
    With ssRst
        If .MaxRows = 0 Then Exit Sub
        .EditEnterAction = EditEnterActionDown
    End With
    fraAccNo.Enabled = False
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

Private Sub ssRst_LostFocus()
    Dim strTmp As String
        
    If ssRst.ActiveRow < 1 Then Exit Sub
    ssRst.Row = ssRst.ActiveRow
    ssRst.Col = 2
    strTmp = ssRst.Value
    ssRst.Row = ssRst.ActiveRow
    ssRst.Col = 18
    If ssRst.Value = "" Then
        ssRst.Col = 2
        ssRst.Value = objPtInfo.GetRstCd(objPtInfo.Result.Item(ssRst.ActiveRow).TestCd, UCase(strTmp))
        ssRst.Col = 18
        ssRst.Value = UCase(strTmp)
    
    End If
    
End Sub

'검사항목별 결과코드 Setting
Private Sub ssRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  '
    If ClickType <> 1 Then Exit Sub

    If MsgFg Then Exit Sub
    MsgFg = True
    If Row <= 0 Then Exit Sub
    objPtInfo.SsTop = picRst.Top + 220
    objPtInfo.SsLeft = picRst.Left - 740
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell
    Call objPtInfo.PopUp(, Col)
    MsgFg = False

  '
End Sub
'결과처리
Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
    
    Dim strRstType As String
    Dim strErr As String
    Dim Col As Long
    Dim Row As Long
   '
    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    On Error GoTo ErrLevaeCell:
   '
    Col = ssRst.ActiveCol
    If Col = objPtInfo.SSCol("RESULT") Then
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
                lblErr.Caption = ""
                objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "결과 입력 오류!"
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
                objPtInfo.NumValCheck
                lblErr.Caption = ""
            End If
        End If
    End If
    
    If Col = objPtInfo.SSCol("RESULT") Then
        Dim strCodeValue As String
        ssRst.Row = Row
        ssRst.Col = 18
        strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            If objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue) <> ssRst.Value Then
                ssRst.Col = 2: ssRst.Value = objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue)
                ssRst.Col = 18: ssRst.Value = strCodeValue
            Else
                ssRst.Col = 18: ssRst.Value = strCodeValue
            End If
        Else
            ssRst.Col = 2: strCodeValue = UCase(Trim(ssRst.Value))
            If objPtInfo.GetRstChk(objPtInfo.Result.Item(Row).TestCd, strCodeValue) = True Then
                ssRst.Col = 2: ssRst.Value = objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue)
                ssRst.Col = 18: ssRst.Value = strCodeValue
            Else
                If strRstType = "F" Then
                    ssRst.Col = 2: ssRst.Value = strCodeValue
                    ssRst.Col = 18: ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = 2: ssRst.Value = strCodeValue
                        ssRst.Col = 18: ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = 2: ssRst.Value = ""
                        ssRst.Col = 18: ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = 2: ssRst.Value = strCodeValue
                    ssRst.Col = 18: ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If

    Exit Sub
   '
ErrLevaeCell:
    lblErr.Caption = strErr
    ssRst.Value = ""
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    DoEvents
   '
    With ssRst
        .Row = Row
        .Col = objPtInfo.SSCol("RESULT")
        .Value = ""
        .Action = ActionActiveCell
        .SetFocus
    End With
    objPtInfo.ResultCheck
    
End Sub
'결과 체크
Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim strRstType As String
    Dim strErr As String
   '
    If Row < 1 Then Exit Sub
    If MsgFg Then Exit Sub

    On Error GoTo ErrLevaeCell
   '
    If Row = ssRst.MaxRows Then
        Call ssRst_LostFocus
        Exit Sub
    End If
    
    lblErr.Caption = ""
    If Col = objPtInfo.SSCol("RESULT") Then
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
    SpDispRtfText NewRow
    
    If Col = objPtInfo.SSCol("RESULT") Then
        Dim strCodeValue As String
        ssRst.Row = Row
        ssRst.Col = 18
        strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            If objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue) <> ssRst.Value Then
                ssRst.Col = 2: ssRst.Value = objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue)
                ssRst.Col = 18: ssRst.Value = strCodeValue
            Else
                ssRst.Col = 18: ssRst.Value = strCodeValue
            End If
        Else
            ssRst.Col = 2: strCodeValue = UCase(Trim(ssRst.Value))
            If objPtInfo.GetRstChk(objPtInfo.Result.Item(Row).TestCd, strCodeValue) = True Then
                ssRst.Col = 2: ssRst.Value = objPtInfo.GetRstCd(objPtInfo.Result.Item(Row).TestCd, strCodeValue)
                ssRst.Col = 18: ssRst.Value = strCodeValue
            Else
                If strRstType = "F" Then
                    ssRst.Col = 2: ssRst.Value = strCodeValue
                    ssRst.Col = 18: ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = 2: ssRst.Value = strCodeValue
                        ssRst.Col = 18: ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = 2: ssRst.Value = ""
                        ssRst.Col = 18: ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = 2: ssRst.Value = strCodeValue
                    ssRst.Col = 18: ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    

    Exit Sub
   '
ErrLevaeCell:
    lblErr.Caption = strErr
    ssRst.EditEnterAction = EditEnterActionDown
   '
    DoEvents
    With ssRst
        .Row = Row
        .Col = objPtInfo.SSCol("RESULT")
        .Value = ""
        .Action = ActionActiveCell
    End With
    objPtInfo.ResultCheck
   '
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    Cancel = True
    ssRst.SetFocus

End Sub

'최근결과보고일시 ToolTip처리
Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
   '
    If Row < 1 Then Exit Sub
    objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
    ssRst.TextTip = TextTipFloatingFocusOnly
   '
End Sub


'소견결과에 따를 Enable=True
Private Sub SpDispRtfText(Optional Row As Long = 0)
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

Private Sub rtfComment_Change()
    If ClearFg Then Exit Sub
End Sub

Private Sub rtfRemark_Change()
    If ClearFg Then Exit Sub
End Sub

Private Sub rtfText_Change()
    If ClearFg Then Exit Sub
End Sub

'접수번호별 결과등록
Private Sub cmdSave_Click()
    
    Dim ii As Long
    Dim blnDBSuccess As Boolean
    Dim strYesNo     As String

    With objPtInfo
        .FootNote = rtfComment.Text
        .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    End With
    '/*
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
            If ssRst.Value = CS_EqpError Then
                ssRst.Action = ActionActiveCell
                Exit Sub
            End If
            If .TxtType = "2" And .RstDiv = "R" Then
                If .TextRst = "" Or ssRst.Value = "" Then
                    '검사는 일반결과와 텍스트 결과를 같이 입력요. 결과보류 처리. (Required 항목일 경우만.. by KMK)
                    ssRst.Col = objPtInfo.SSCol("EC")
                    ssRst.Value = 1
                End If
            End If
            ssRst.Col = objPtInfo.SSCol("RESULT")
        End With
    Next ii
   '
    blnDBSuccess = objPtInfo.DataEntry 'objPtInfo                  '결과등록을 수행한다.
    If blnDBSuccess = False Then
        Call ClearData
        MsgBox objPtInfo.ErrNo & _
                " - " & objPtInfo.ErrText, vbCritical + vbOKOnly, "결과등록 ERROR"
        Exit Sub
    Else
        If objDic.Exists(mskAccNo.Tag) Then
            objDic.KeyChange mskAccNo.Tag
            objDic.Fields("donefg") = "1"
        End If
        Call ClearData
    End If

    ssRst.MaxRows = 0
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""

End Sub

'접수번호리스트 Visible=True
Private Sub cmdAcc_Click()
    Dim itmFound As ListItem
    
    lvwCount.ListItems.Clear
    If objDic.RecordCount > 0 Then
        With lvwCount
            objDic.MoveFirst
            Do Until objDic.EOF
                Set itmFound = .ListItems.Add(, , objDic.Fields("seq"))
                itmFound.SubItems(1) = Trim(objDic.Fields("workarea") & "-" & _
                                            objDic.Fields("accdt") & "-" & _
                                            objDic.Fields("accseq"))
                itmFound.SubItems(2) = IIf(objDic.Fields("donefg") = "1", "확인", "미확")
                objDic.MoveNext
            Loop
        End With
    End If
    fraAcc.Visible = True
End Sub
'접수번호리스트 Visible=False
Private Sub cmdCloseFra_Click()
    fraAcc.Visible = False
End Sub

