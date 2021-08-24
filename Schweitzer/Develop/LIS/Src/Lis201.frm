VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm201WSBuild 
   BackColor       =   &H00DBE6E6&
   Caption         =   "워크쉬트 생성"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   14535
   Tag             =   "20100"
   WindowState     =   2  '최대화
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   4140
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "엑셀저장(&E)"
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
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
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
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraBuild 
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ Worksheet 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   75
      TabIndex        =   11
      Top             =   2325
      Width           =   3225
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   315
         Index           =   1
         Left            =   3930
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   405
         Width           =   705
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "개별"
         Height          =   315
         Index           =   0
         Left            =   3225
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   405
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   13125
         Picture         =   "Lis201.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   49
         Tag             =   "121"
         Top             =   705
         Width           =   615
      End
      Begin VB.Frame fraMinMax 
         BorderStyle     =   0  '없음
         Height          =   255
         Left            =   2640
         TabIndex        =   44
         Top             =   105
         Width           =   525
         Begin VB.CommandButton cmdMax 
            Caption         =   "□"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   270
            TabIndex        =   46
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton cmdMin 
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   15
            TabIndex        =   45
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CheckBox chkFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수시간별 순서로 정렬"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   135
         TabIndex        =   35
         Top             =   390
         Width           =   2220
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
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
         Height          =   510
         Left            =   1800
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   5625
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvwBuildList 
         Height          =   4905
         Left            =   105
         TabIndex        =   9
         Top             =   705
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "작업번호"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "접수번호"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwPatient2 
         Height          =   600
         Left            =   3225
         TabIndex        =   47
         Tag             =   "20113"
         Top             =   720
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   1058
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin MSComctlLib.ListView lvwLabNum 
         Height          =   4410
         Left            =   3225
         TabIndex        =   48
         Tag             =   "20113"
         Top             =   1350
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   7779
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
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
         Left            =   2520
         MouseIcon       =   "Lis201.frx":0102
         MousePointer    =   99  '사용자 정의
         TabIndex        =   30
         Top             =   405
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   285
         Index           =   1
         Left            =   2430
         Shape           =   4  '둥근 사각형
         Top             =   375
         Width           =   675
      End
   End
   Begin VB.Frame fraAdd 
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ 추가 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   3300
      TabIndex        =   12
      Top             =   2325
      Width           =   11160
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00F4F0F2&
         Caption         =   "&Add"
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
         Height          =   600
         Left            =   10050
         Picture         =   "Lis201.frx":040C
         Style           =   1  '그래픽
         TabIndex        =   7
         Tag             =   "121"
         Top             =   720
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwPatient 
         Height          =   600
         Left            =   240
         TabIndex        =   8
         Tag             =   "20113"
         Top             =   735
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   1058
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin MSComctlLib.ListView lvwAddList 
         Height          =   4785
         Left            =   240
         TabIndex        =   10
         Tag             =   "20113"
         Top             =   1335
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   8440
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
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
      Begin MSMask.MaskEdBox mskAccNo 
         Height          =   330
         Left            =   1680
         TabIndex        =   5
         Top             =   375
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
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
      Begin MSMask.MaskEdBox mskAddSeq 
         Height          =   330
         Left            =   4695
         TabIndex        =   6
         Top             =   375
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
         AutoTab         =   -1  'True
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
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   255
         TabIndex        =   52
         Top             =   375
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "추가 접수번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   3630
         TabIndex        =   57
         Top             =   375
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "Work No."
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "☞ 입력하지 않고 추가하면 번호가 자동부여 됩니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C76456&
         Height          =   240
         Left            =   5385
         TabIndex        =   34
         Top             =   465
         Width           =   5025
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
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
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
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
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   8310
      TabIndex        =   36
      Top             =   -150
      Visible         =   0   'False
      Width           =   3465
      Begin DRcontrol1.DrFrame fraDatePicker 
         Height          =   450
         Left            =   1515
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         Appearance      =   0
         Title           =   ""
         DelLine         =   0
         BackColor       =   14411494
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComCtl2.DTPicker txtDatePicker 
            Height          =   360
            Left            =   45
            TabIndex        =   38
            Top             =   15
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   635
            _Version        =   393216
            CustomFormat    =   "yyy-MM-dd HH:mm"
            Format          =   96010243
            CurrentDate     =   36328
         End
      End
      Begin VB.Label lblModifyLast 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "◈ 최근 마감시간 :"
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
         Left            =   0
         MouseIcon       =   "Lis201.frx":050E
         MousePointer    =   99  '사용자 정의
         TabIndex        =   40
         Tag             =   "20102"
         ToolTipText     =   "Click하시면 마감시간을 수정할 수 있습니다."
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblLastDtTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "1999/01/01 12:30"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1785
         TabIndex        =   39
         Tag             =   "20102"
         Top             =   300
         Width           =   1650
      End
   End
   Begin VB.OptionButton optCondition 
      BackColor       =   &H00EAE7E3&
      Caption         =   "&Edit"
      Height          =   480
      Index           =   1
      Left            =   1395
      Style           =   1  '그래픽
      TabIndex        =   29
      Top             =   45
      Width           =   1320
   End
   Begin VB.OptionButton optCondition 
      BackColor       =   &H00EAE7E3&
      Caption         =   "&New"
      Height          =   480
      Index           =   0
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   45
      Width           =   1320
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ WorkSheet 정보"
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
      Height          =   1845
      Left            =   8325
      TabIndex        =   1
      Top             =   450
      Width           =   6135
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   5
         Left            =   75
         TabIndex        =   58
         Top             =   450
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "Build Count"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   7
         Left            =   75
         TabIndex        =   59
         Top             =   810
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "Add Count "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   8
         Left            =   75
         TabIndex        =   60
         Top             =   1170
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "Total Build Count"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   9
         Left            =   2805
         TabIndex        =   61
         Top             =   450
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "Work No."
         Appearance      =   0
      End
      Begin VB.Label lblCntTotal 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1965
         TabIndex        =   18
         Top             =   1170
         Width           =   645
      End
      Begin VB.Label lblToSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4755
         TabIndex        =   17
         Top             =   450
         Width           =   645
      End
      Begin VB.Label lblFrSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3870
         TabIndex        =   16
         Top             =   450
         Width           =   645
      End
      Begin VB.Label lblCntAdd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1965
         TabIndex        =   15
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblCntBuild 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00D1D8D3&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1965
         TabIndex        =   14
         Top             =   450
         Width           =   645
      End
      Begin VB.Line Line1 
         X1              =   4560
         X2              =   4690
         Y1              =   570
         Y2              =   570
      End
   End
   Begin VB.Frame fraWSHeader 
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
      Height          =   1845
      Left            =   75
      TabIndex        =   13
      Top             =   450
      Width           =   8235
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
         Left            =   6390
         Style           =   2  '드롭다운 목록
         TabIndex        =   64
         Top             =   270
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CheckBox chkSpcNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체번호로 작성"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006B72A9&
         Height          =   315
         Left            =   5985
         TabIndex        =   43
         Top             =   1470
         Width           =   1740
      End
      Begin VB.OptionButton optStatFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009A617E&
         Height          =   225
         Index           =   2
         Left            =   5100
         TabIndex        =   33
         Top             =   1140
         Width           =   705
      End
      Begin VB.OptionButton optStatFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "비응급"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009A617E&
         Height          =   225
         Index           =   1
         Left            =   4095
         TabIndex        =   32
         Top             =   1140
         Width           =   975
      End
      Begin VB.OptionButton optStatFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009A617E&
         Height          =   225
         Index           =   0
         Left            =   3180
         TabIndex        =   31
         Top             =   1140
         Width           =   690
      End
      Begin VB.TextBox txtToSeq 
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
         Left            =   3690
         MaxLength       =   8
         TabIndex        =   27
         Top             =   1035
         Width           =   1125
      End
      Begin VB.TextBox txtFromSeq 
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
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   26
         Top             =   1035
         Width           =   1125
      End
      Begin VB.CommandButton cmdWSList 
         BackColor       =   &H00DEDBDD&
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
         Height          =   315
         Left            =   3405
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   270
         Width           =   270
      End
      Begin VB.CommandButton cmdBuild 
         BackColor       =   &H00FFF7FC&
         Caption         =   "조회 (&Q)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6000
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   1035
         Width           =   1275
      End
      Begin VB.TextBox txtWorkCd 
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
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   0
         Top             =   270
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpWorkDt 
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   675
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
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
         Format          =   96010240
         CurrentDate     =   36318
      End
      Begin MSComCtl2.DTPicker dtpWorkTm 
         Height          =   330
         Left            =   5295
         TabIndex        =   3
         Top             =   675
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
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
         Format          =   96010242
         CurrentDate     =   -30
      End
      Begin MedControls1.LisLabel lblWorkCdNm 
         Height          =   315
         Left            =   3690
         TabIndex        =   22
         Top             =   270
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         BackColor       =   15726072
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
      Begin MSMask.MaskEdBox mskSpcNo 
         Height          =   330
         Left            =   2280
         TabIndex        =   41
         Top             =   1410
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "############"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   53
         Top             =   270
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "WorkSheet Code "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDateLabel 
         Height          =   315
         Left            =   210
         TabIndex        =   54
         Top             =   660
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "기준 접수일시 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWorkNoLabel 
         Height          =   315
         Left            =   210
         TabIndex        =   55
         Top             =   1035
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "Work No "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNo 
         Height          =   315
         Left            =   210
         TabIndex        =   56
         Top             =   1410
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "검체번호 "
         Appearance      =   0
      End
      Begin VB.Label lblLabNo 
         BackStyle       =   0  '투명
         Caption         =   "Label3"
         Height          =   300
         Left            =   3915
         TabIndex        =   42
         Top             =   1440
         Width           =   2070
      End
      Begin VB.Line Line2 
         X1              =   3465
         X2              =   3595
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblWorkArea 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Height          =   315
         Left            =   6630
         TabIndex        =   24
         Top             =   315
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin VB.ListBox lstWSCode 
      Appearance      =   0  '평면
      BackColor       =   &H00F4FDFF&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   2355
      TabIndex        =   21
      Top             =   1095
      Visible         =   0   'False
      Width           =   4980
   End
End
Attribute VB_Name = "frm201WSBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objLab032       As clsComcode032
Private objLab301       As clsWSBuild
Private objPtInfo       As clsPatientInfo
Private objDeleteList   As New clsDictionary

Private blnFirst        As Boolean
Private blnDayCount     As Boolean
Private blnChange       As Boolean
Private gblnNewObj      As Boolean

Private gstrPtAddInfo   As String
Private gstrMsk         As String
Private gstrLastDt      As String
Private gstrLastTm      As String
Private intCurFrame     As Integer      ' 현재 프레임

Private Sub chkFg_Click()
    
    lvwBuildList.ListItems.Clear
    lvwPatient2.ListItems.Clear
    lvwLabNum.ListItems.Clear
    optDiv(0).Value = True
    Call cmdBuild_Click
End Sub

Private Sub chkSpcNo_Click()
    lvwBuildList.ListItems.Clear
    If chkSpcNo.Value = 0 Then
        lblSpcNo.Visible = False
        mskSpcNo.Visible = False
        lblLabNo.Visible = False
        cmdBuild.Enabled = True
    Else
        cmdBuild.Enabled = False
        lblSpcNo.Visible = True
        mskSpcNo.Visible = True
        lblLabNo.Visible = False
        lblLabNo.Caption = ""
        mskSpcNo.Text = "____________"
        mskSpcNo.SetFocus
    End If
    Call cmdMin_Click
    
End Sub

Private Sub cmdBuild_Click()
    
    Dim strTm As String
    Dim aryTmp() As String
    Dim DataExists As Boolean
    Dim objProBar As clsProgress
    Dim strStatFg As String
    Dim strChkfg As String
    
    
    strChkfg = IIf(chkFg.Value = "1", "1", "2")
    
    '
    If txtWorkCd.Text = "" Then Exit Sub
    '
    cmdBuild.Enabled = False
    cmdClear.Enabled = False
    cmdExit.Enabled = False

    MouseRunning
    
    Set objProBar = New clsProgress
    With objProBar
        .Container = Me
        .Left = optCondition(1).Left + optCondition(1).Width + 20
        .Top = optCondition(1).Top + optCondition(1).Height - 390
        .Width = fraWSHeader.Width - (optCondition(1).Width * 2)
        .Height = 370
        .Message = "WORKSHEET 대상을 검색 중입니다..."
        .Max = 90
        .Value = 10
        DoEvents
    End With
    
    DoEvents

    With objLab301

        DoEvents
        
        strStatFg = IIf(optStatFg(0).Value, "1", IIf(optStatFg(1).Value, "0", "2"))
        
'        DataExists = .LoadTable_NEW(txtWorkCd.Text, DateStr(dtpWorkDt.Value), _
                                  Format(dtpWorkTm.Value, "hhmmss"), ObjMyUser.EmpId, ObjSysInfo.BuildingCd, _
                                  DateStr(dtpWorkDt.Value) & Format(dtpWorkTm.Value, "hhmmss"), lblWorkarea.Caption, optCondition(0).Value, _
                                  txtFromSeq.Text, txtToSeq.Text, strStatFg, strChkfg)
                                  
        '2014-01-27 병동코드추가삽입
        DataExists = .LoadTable_NEW_2014(txtWorkCd.Text, DateStr(dtpWorkDt.Value), _
                                  Format(dtpWorkTm.Value, "hhmmss"), ObjMyUser.EmpId, cboPoct.ListIndex, ObjSysInfo.BuildingCd, _
                                  DateStr(dtpWorkDt.Value) & Format(dtpWorkTm.Value, "hhmmss"), lblWorkarea.Caption, optCondition(0).Value, _
                                  txtFromSeq.Text, txtToSeq.Text, strStatFg, strChkfg)
                                  
        objProBar.Value = 50
        If DataExists Then
            objProBar.Value = 70
            medDataLoadLvw lvwBuildList, vbNewLine, vbTab, .GetStrBuildList
            objProBar.Value = 90
        Else
            MouseDefault
            MsgBox "해당 데이타가 없습니다"
            Call cmdClear_Click
            Exit Sub
        End If
    End With
    '
    EditData
    DisplayCount
    mskAccNo.SetFocus
    DoEvents

    MouseDefault
    Set objProBar = Nothing

    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
    If optCondition(0).Value Then
        blnChange = True
    Else
        blnChange = False
    End If
   '
End Sub

Private Sub cmdClear_Click()
    ClearData
    txtWorkCd.SetFocus
    blnChange = False
    Call cmdMin_Click
End Sub

Private Sub cmdDelete_Click()
    
    Dim i As Long
    Dim strKey As String, strData As String
    Dim Resp As VbMsgBoxResult
    
    Resp = MsgBox("선택된 접수번호를 해당 Worksheet에서 삭제하시겠습니까?", vbQuestion + vbYesNo, "Worksheet작성")
    If Resp = vbNo Then Exit Sub
    
    With lvwBuildList
        For i = .ListItems.Count To 1 Step -1
            If .ListItems(i).Checked Then
                strKey = .ListItems(i).Text
                strData = .ListItems(i).SubItems(1)
                If Not objDeleteList.Exists(strKey) Then objDeleteList.AddNew strKey, strData
                .ListItems.Remove i
            End If
        Next
    End With
    
    blnChange = True
    
End Sub

Private Sub cmdExcel_Click()
    Dim OneRec As String '그리드 한행의 내용을 가지고 있는 변수
    Dim FileName As String
    Dim rCount As Integer
    Dim cCount As Integer
    Dim F As String
    Dim i As Integer
    Dim j As Integer
   
    'Cancel을 True로 설정한다.
    DlgSave.CancelError = True
On Error GoTo Err_Handler:
    OneRec = ""
    FileName = ""
    'Flags상수 : cdlOFNOverwritePrompt(이미 존재하는 화일을 선택한 경우 에러처리)
    '           cdlOFNExplorer(탐색기와 같은 형태의 파일선택 화면 (Win95, 32bit))
    '           cdlOFNLongNames(긴 파일 이름(Long File Name) 허용)
    DlgSave.Flags = cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames
    DlgSave.Filter = "엑셀파일 (*.xls) |*.xls|모든파일 (*.*)|*.*"
    DlgSave.DialogTitle = "엑셀화일 형태로만 저장됩니다"
    DlgSave.InitDir = App.Path
    DlgSave.FileName = FileName
    DlgSave.ShowSave
    
    If Len(DlgSave.FileName) = 0 Then Exit Sub
    
    F = FreeFile()
    rCount = lvwLabNum.ListItems.Count
    cCount = lvwLabNum.ListItems.Item(1).ListSubItems.Count
    
    Open DlgSave.FileName For Output As #F
        For i = 1 To rCount
            OneRec = lvwLabNum.ListItems.Item(i) & vbTab
            For j = 1 To cCount
                OneRec = OneRec & lvwLabNum.ListItems.Item(i).SubItems(j) & vbTab
            Next j
            Print #F, OneRec
            OneRec = ""
        Next i
    Close #F

    Exit Sub
    
Err_Handler:
    If Err.Number = cdlCancel Then
        '취소단추를 눌렀습니다.
    Else
        MsgBox Err.Number & ":" & Err.Description, vbQuestion
    End If
    Exit Sub

End Sub

Private Sub cmdExit_Click()
    
    Dim Resp As VbMsgBoxResult
    
    If blnChange Then
        Resp = MsgBox("변경된 데이타를 저장하지 않고 종료하시겠습니까?", vbQuestion + vbYesNo, "Worksheet작성")
        If Resp = vbNo Then Exit Sub
    End If
    
    Set objLab301 = Nothing
    Set objLab032 = Nothing
    Set objPtInfo = Nothing
    Set objDeleteList = Nothing
    Unload Me
    Set frm201WSBuild = Nothing
    
End Sub

Private Sub cmdMax_Click()
    Dim ii          As Integer
    Dim strTmp      As String
    Dim strFirst    As String
    
    cmdSort.Enabled = True
    
    With lvwBuildList
        If .ListItems.Count < 1 Then Exit Sub
        strFirst = .ListItems.Item(1).Text
        For ii = 1 To .ListItems.Count
            strTmp = .ListItems.Item(ii).Text
            If strTmp <> strFirst Then
                cmdSort.Enabled = False
                Exit Sub
            End If
            strFirst = Val(strFirst) + 1
        Next
    End With
    
    fraBuild.Width = 14385
    fraMinMax.Left = 13830
    fraBuild.ZOrder 0
End Sub

Private Sub cmdMin_Click()

    
    fraBuild.Width = 3225
    
    fraMinMax.Left = 2640
    fraBuild.ZOrder 0
    lvwPatient2.ListItems.Clear
    lvwLabNum.ListItems.Clear
    optDiv(0).Value = True
    
End Sub

Private Sub cmdSave_Click()
    
    Dim objProBar     As clsProgress
    Dim Resp          As VbMsgBoxResult
    
    Dim blnResp       As Boolean
    
    Dim strTmp        As String
    Dim strAddAccNo   As String
    Dim strDelWorkSeq As String
    
    Dim intStartWSSeq As Integer
    Dim ii            As Integer
    
    Dim DelWorkDt   As String
    Dim DelWorkCd   As String
    
    '

    If lvwAddList.ListItems.Count = 0 And lvwBuildList.ListItems.Count = 0 Then Exit Sub
   '
    cmdSave.Enabled = False

    MouseRunning
    
    Set objProBar = New clsProgress
    With objProBar
        .Container = Me
        .Left = fraAdd.Left
        .Top = fraAdd.Top
        .Width = fraAdd.Width  ' (tabView.Width - 1700)
        .Height = 260
        .Message = "WORKSHEET 내역을 저장 중입니다..."
        .Max = 90
        
'        .SetMyForm Me
'        .Choice = True
'        .XPos = fraAdd.Left
'        .YPos = fraAdd.Top
'        .XWidth = fraAdd.Width  ' (tabView.Width - 1700)
'        .ForeColor = &H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = 260
'        .MSG = "Worksheet 내역을 저장 중입니다...."
'        .Max = 90
'        .Value = 1
        DoEvents
    End With
    
    DoEvents

'==========================================================================================
'[빌드된 리스트 뷰를 읽어서 WorkSheet를 생성한다.]
'작성일  : 2003.12.22
'작성자  : 김정규
'작성이유: WorkSheet마스터에서 검사발생순위가 높은순으로 Sort기능 필요로
'==========================================================================================
    Dim strBuildList As String
    Dim objWSSave    As clsLisSqlResult

    'Worksheet 편집모드 - Delete
    If optCondition(1).Value Then
        If objDeleteList.RecordCount > 0 Then
            strDelWorkSeq = ""
            objDeleteList.MoveFirst
            While Not objDeleteList.EOF
                strDelWorkSeq = strDelWorkSeq & objDeleteList.Key & vbTab
                objDeleteList.MoveNext
            Wend
        End If
        blnResp = objLab301.Delete(strDelWorkSeq, txtWorkCd.Text, Format(dtpWorkDt.Value, CS_DateDbFormat))
        If Not blnResp Then
            MsgBox "선택된 검체 삭제 중 오류가 발생했습니다. 전산실 혹은 임상병리과로 연락바랍니다.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbCritical, "오류"
            Exit Sub
        End If
    End If

    Set objWSSave = New clsLisSqlResult

    '[WokrSheet]신규 작성 모드
    If lvwBuildList.ListItems.Count > 0 And optCondition(0).Value Then  '신규
        For ii = 1 To lvwBuildList.ListItems.Count
            If ii = 1 Then intStartWSSeq = lvwBuildList.ListItems.Item(ii).Text
            strTmp = DBAccNo(lvwBuildList.ListItems.Item(ii).SubItems(1)) & COL_DIV & "" & COL_DIV & ""
            If strBuildList <> "" Then
                strBuildList = strBuildList & vbTab & strTmp
            Else
                strBuildList = strTmp
            End If
        Next
    Else
    '[WokrSheet]EDIT 작성 모드, mskadd.clip가 "" 인경우 intStartWSSeq  를 다시 구한다.
        strBuildList = ""
        If lvwAddList.ListItems.Count > 0 Then
            If lvwAddList.ListItems.Item(1).Text <> "" Then
                intStartWSSeq = lvwAddList.ListItems.Item(1).Text
            Else
                intStartWSSeq = Val(mskAddSeq.ClipText)
                If intStartWSSeq = 0 Then
                    intStartWSSeq = CInt(objLab301.StartWorkSeq(txtWorkCd.Text, Format(dtpWorkDt.Value, CS_DateDbFormat)))
                End If
            End If

            For ii = 1 To lvwAddList.ListItems.Count
                strTmp = DBAccNo(lvwAddList.ListItems.Item(ii).SubItems(1)) & COL_DIV & lvwAddList.ListItems.Item(ii).Text & COL_DIV & ""
                If strBuildList <> "" Then
                    strBuildList = strBuildList & vbTab & strTmp
                Else
                    strBuildList = strTmp
                End If
            Next ii
        End If
    End If

    With objWSSave
        If strBuildList <> "" Then
            Call .BuildWorkSheet(Format(dtpWorkDt.Value, CS_DateDbFormat), txtWorkCd.Text, _
                                 Format(Time, "hhmmss"), CLng(intStartWSSeq), ObjSysInfo.EmpId, strBuildList, _
                                 objLab301.GetLastDt, objLab301.GetLastTm, objProBar)
            If Not .SaveOk = True Then
                MsgBox "WorkSheet 작성도중 Error가 발생하였습니다.", vbInformation + vbOKOnly, "Info"
            Else
                Resp = MsgBox("지금 출력하시겠습니까 ? ", vbYesNo, "워크쉬트 출력")
                If Resp = vbYes Then
                    objProBar.Message = "Worksheet List를 츨력 중입니다... 잠시만 기다리세요! "
                    DoEvents
                    Call PrintWorkList(txtWorkCd.Text, lblWorkCdNm.Caption, lblFrSeq.Caption, lblToSeq.Caption)
                End If
            End If
        End If
    End With

    objProBar.Message = "4개월 전의 Worksheet 내역을 삭제 중입니다...."
    DoEvents

    DelWorkCd = txtWorkCd.Text
    DelWorkDt = Format(DateAdd("d", -120, dtpWorkDt.Value), "yyyymmdd")

    Call objLab301.WorkHistoryDelete(DelWorkCd, DelWorkDt)

    MouseDefault
    Set objProBar = Nothing
    Set objWSSave = Nothing
    Call ClearData: Call cmdMin_Click
    
    Exit Sub
'==========================================================================================

    'strAddAccNo : 추가검체의 접수번호 & 지정된 WorkNo & 추가여부
    strAddAccNo = ""
    If lvwAddList.ListItems.Count > 0 Then
        For ii = 1 To lvwAddList.ListItems.Count
            If ii = 1 Then
                strAddAccNo = lvwAddList.ListItems(ii).SubItems(1) & COL_DIV & lvwAddList.ListItems(ii).Text & COL_DIV & "Y"
            Else
                strAddAccNo = strAddAccNo & vbTab & lvwAddList.ListItems(ii).SubItems(1) & COL_DIV & _
                                                    lvwAddList.ListItems(ii).Text & COL_DIV & "Y"
            End If
        Next ii
    End If
    '
    If optCondition(1).Value Then   'Worksheet 편집 - Delete
        If objDeleteList.RecordCount > 0 Then
            strDelWorkSeq = ""
            objDeleteList.MoveFirst
            While Not objDeleteList.EOF
                strDelWorkSeq = strDelWorkSeq & objDeleteList.Key & vbTab
                objDeleteList.MoveNext
            Wend
        End If
        blnResp = objLab301.Delete(strDelWorkSeq, txtWorkCd.Text, Format(dtpWorkDt.Value, CS_DateDbFormat))
        If Not blnResp Then
            MsgBox "선택된 검체 삭제 중 오류가 발생했습니다. 전산실 혹은 임상병리과로 연락바랍니다.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbCritical, "오류"
            Exit Sub
        End If
    Else
        If objDeleteList.RecordCount > 0 Then
            objDeleteList.MoveFirst
            While Not objDeleteList.EOF
                strDelWorkSeq = Format(dtpWorkDt.Value, CS_DateDbFormat) & Trim(txtWorkCd.Text) & objDeleteList.Key
                objLab301.RemoveItem (strDelWorkSeq)
                objDeleteList.MoveNext
            Wend
        End If
    End If
    
    If lvwBuildList.ListItems.Count > 0 And optCondition(0).Value Then  '신규
        blnResp = objLab301.Insert(Format(Time, "hhmmss"), strAddAccNo, , , , objProBar)
    Else
       '추가내역만 BUILD하는 경우
        blnResp = objLab301.Insert(Format(Time, "hhmmss"), strAddAccNo, txtWorkCd.Text, _
                                   DateStr(dtpWorkDt.Value), ObjMyUser.EmpId, objProBar, False)
    End If
    If Not blnResp Then
        MsgBox "Worksheet 저장 중 오류가 발생했습니다. 전산실 혹은 임상병리과로 연락바랍니다.(☎" & ObjSysInfo.HelpLine & ")", _
                vbCritical, "오류"
        Exit Sub
    End If
    '
    'Worksheet작성이 성공한 경우만...
    If blnResp Then
        With objLab301
            If .ErrText <> "" Then
                MsgBox .ErrNo & " - " & .ErrText, vbCritical + vbOKOnly, "워크쉬트 생성 ERROR"
            Else
                Resp = MsgBox("지금 출력하시겠습니까 ? ", vbYesNo, "워크쉬트 출력")
                If Resp = vbYes Then
                    objProBar.Message = "Worksheet List를 츨력 중입니다... 잠시만 기다리세요! "
                    DoEvents
                    Call PrintWorkList(txtWorkCd.Text, lblWorkCdNm.Caption, lblFrSeq.Caption, lblToSeq.Caption)
                    
                    '-- 추가 업무나열서
                    
                End If
            End If
        End With
    End If
    '
    
    '이미 작성된 WorkSheet내역을 삭제한다.
    '4개월 이전의 데이타에 대해서만 삭제한다.
    'kjg(최부장님 요청)
    
    objProBar.Message = "4개월 전의 Worksheet 내역을 삭제 중입니다...."
    DoEvents
    
    DelWorkCd = txtWorkCd.Text
    DelWorkDt = Format(DateAdd("d", -120, dtpWorkDt.Value), "yyyymmdd")
    
    Call objLab301.WorkHistoryDelete(DelWorkCd, DelWorkDt)
    
    MouseDefault
    Set objProBar = Nothing

    ClearData
    '

End Sub

Private Sub PrintWorkList(ByVal pWorkCd As String, ByVal pWorkNm As String, _
                                    ByVal pFrSeq As String, ByVal pToSeq As String)

    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        Exit Sub
    End If
    
    Dim MyReport As New clsWorkListG
    
    With MyReport
        .WorkCode = pWorkCd
        .WorkName = pWorkNm
        .WorkDate = Format(Now, CS_DateDbFormat)
        .FromSeq = pFrSeq
        .ToSeq = pToSeq
        Call .Print_Worksheet
    End With

    Set MyReport = Nothing
    Exit Sub

End Sub




Private Sub cmdWSList_Click()

    If lstWSCode.ListCount = 0 Then
        MsgBox "등록된 Worksheet 코드가 없습니다.", vbExclamation, "메세지"
        Exit Sub
    End If
    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
    Call medCodeHelp(0, lstWSCode, txtWorkCd.Text, txtWorkCd, dtpWorkDt)

End Sub

Private Sub dtpWorkDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpWorkTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    
    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then lstWSCode.Visible = False
End Sub

Private Sub Form_Load()
   '
    Me.Show
    DoEvents
    
    chkFg.Value = 1
    
    Set objPtInfo = New clsPatientInfo
    Call objPtInfo.LoadWorkSheetCode(ObjSysInfo.BuildingCd, lstWSCode)
    Call LoadLvwHead
    DoEvents
    Call ClearData
    '
    
    cboPoct.Clear
    cboPoct.AddItem "0.전체"
    cboPoct.AddItem "1.재활센타"
    cboPoct.AddItem "2.본관"
    KeyPreview = True
    blnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objLab301 = Nothing
    Set objLab032 = Nothing
    Set objPtInfo = Nothing
    Set objDeleteList = Nothing
End Sub

Private Sub lblModifyLast_Click()
    If lblLastDtTm.Caption = "" Then Exit Sub
    txtDatePicker.Value = lblLastDtTm.Caption
    fraDatePicker.Visible = True
    fraDatePicker.ZOrder 0
    txtDatePicker.SetFocus
End Sub

Private Sub lblReset_Click()
    Dim i As Long
    For i = 1 To lvwBuildList.ListItems.Count
        'lvwBuildList.ListItems(i).SELECTed = False
        lvwBuildList.ListItems(i).Checked = False
    Next
End Sub

Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyReturn:
        txtWorkCd.Text = medGetP(lstWSCode.Text, 1, vbTab)
        lblWorkCdNm.Caption = medGetP(lstWSCode.Text, 2, vbTab)
        lblWorkarea.Caption = medGetP(lstWSCode.Text, 3, vbTab)
        lstWSCode.Visible = False
        dtpWorkDt.SetFocus
        Call txtWorkCd_Validate(False)
        
        Select Case txtWorkCd.Text
         Case "POCT" '현장검사
              cboPoct.Visible = True
         Case Else
              cboPoct.Visible = False
        End Select
        cboPoct.ListIndex = "0"
    End Select

End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstWSCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lstWSCode.SetFocus
End Sub

Private Sub lvwLabNum_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvwPatient2.ListItems.Clear
    With objPtInfo
        .PtType = RESULT_BY_DEFAULT
        .AccNo = AccTrim(Item.SubItems(1))
        .LoadTable
        If .RecordCount > 0 Then
            medDataLoadLvw lvwPatient2, vbNewLine, vbTab, .GetStringPtInfo
        End If
    End With
End Sub

Private Sub mskAccNo_Change()
    mskAddSeq.Text = "____"
End Sub

Private Sub mskAccNo_GotFocus()
    mskAccNo.SelStart = 0
End Sub

Private Sub mskAddSeq_GotFocus()
    mskAddSeq.SelStart = 0
End Sub

Private Sub mskAddSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call mskAddSeq_Validate(False)
End Sub

Private Sub mskAddSeq_Validate(Cancel As Boolean)
    
    If mskAddSeq.ClipText = "" Then Exit Sub
    
    Dim strAddSeq   As String
    Dim lngLastSeq  As Long
    Dim itmFound    As ListItem
    
    lngLastSeq = objLab301.StartWorkSeq(txtWorkCd.Text, Format(dtpWorkDt.Value, CS_DateDbFormat))
    
    strAddSeq = mskAddSeq.ClipText
    If Val(strAddSeq) >= lngLastSeq Then
        MsgBox "재사용 할 수 없는 Work Seq 입니다.", vbExclamation, "검체추가"
        Cancel = True
        mskAddSeq.Text = "____"
        mskAddSeq.SetFocus
        Exit Sub
    ElseIf Val(strAddSeq) >= Val(lblFrSeq.Caption) And Val(strAddSeq) <= Val(lblToSeq.Caption) Then
    '조회된 Worksheet List에서 Check...
        Set itmFound = lvwBuildList.FindItem(strAddSeq, lvwText, , 0)   'lvwWholeWord)
        If itmFound Is Nothing Then
            cmdAdd.SetFocus
        Else
            itmFound.EnsureVisible
            MsgBox "이미 작성된 Work Seq 입니다.", vbExclamation, "검체추가"
            Cancel = True
            mskAddSeq.Text = "____"
            mskAddSeq.SetFocus
            Exit Sub
        End If
    Else
    'Database의 저장된 Worksheet내역 Check...
        If objPtInfo.IsWorkSeqExists(txtWorkCd.Text, _
                                     Format(dtpWorkDt.Value, CS_DateDbFormat), _
                                     mskAddSeq.ClipText) Then
            MsgBox "이미 작성된 Work Seq 입니다.", vbExclamation, "검체추가"
            Cancel = True
            mskAddSeq.Text = "____"
            mskAddSeq.SetFocus
            Exit Sub
        Else
            cmdAdd.SetFocus
        End If
    End If
    
End Sub

Private Sub optCondition_Click(Index As Integer)
    
    optCondition(Index).ForeColor = vbBlue
    optCondition((Index + 1) Mod 2).ForeColor = vbBlack
'
    If Index = 0 Then
        lblDateLabel.Caption = "기준 접수일시 : "
        dtpWorkTm.Visible = True
        lblWorkNoLabel.Visible = False
        txtFromSeq.Visible = False
        txtToSeq.Visible = False
        optStatFg(0).Visible = True
        optStatFg(1).Visible = True
        optStatFg(2).Visible = True
        chkSpcNo.Enabled = True
    Else
        lblDateLabel.Caption = "작업일자 : "
        dtpWorkTm.Visible = False
        lblWorkNoLabel.Visible = True
        txtFromSeq.Visible = True
        txtToSeq.Visible = True
        optStatFg(0).Visible = False
        optStatFg(1).Visible = False
        optStatFg(2).Visible = False
        chkSpcNo.Enabled = False
    End If
    Call ClearData
    Call cmdMin_Click
    If txtWorkCd.Enabled Then txtWorkCd.SetFocus
End Sub

'최근 마감시간 변경...
Private Sub txtDatePicker_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lblLastDtTm.Caption = Format(txtDatePicker.Value, CS_DateLongFormat & " " & CS_TimeShortFormat)
        gstrLastDt = Format(Format(txtDatePicker.Value, CS_DateLongFormat), CS_DateDbFormat)
        gstrLastTm = Format(Format(txtDatePicker.Value, CS_TimeLongFormat), CS_TimeDbFormat)
        dtpWorkDt.SetFocus
        fraDatePicker.Visible = False
    End If
End Sub

Private Sub lvwAddList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwAddList.SortKey = ColumnHeader.Index - 1
    lvwAddList.Sorted = True
End Sub

Private Sub lvwAddList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Label6.Caption = LvwClickData(Item)
End Sub

Private Sub LoadLvwHead()
    
    Dim colHead As ColumnHeader
    Dim intMode As Integer
    
    '국가별 설정 모드
    intMode = 1          'Korea
    'intMode = 2         'English
    If intMode = 1 Then
'        medInitLvwHead lvwPatient2, "환자 ID,성명,성/나이,생년월일,병상,주치의,검체, 접수일자", _
'                                   "-100,300,-400,0,100,100,0"  ' "-200,500,-500,0,100,100,0"
        medInitLvwHead lvwPatient2, "환자 ID,성명,성/나이,생년월일,병상,주치의,검체, 접수일자,비고(외부QC)", _
                                   "-100,300,-400,0,100,100,100,0"  ' "-200,500,-500,0,100,100,0"
    
        medInitLvwHead lvwLabNum, "번호,접수 번호,환자 ID,성 명,검 체,검 사,비 고", _
                                   "-700,0,-400,-100,100,3000,0"
                                   
        medInitLvwHead lvwPatient, "환자 ID,성명,성/나이,생년월일,병상,주치의,검체, 접수일자,비고(외부QC)", _
                                   "-100,300,-400,0,100,100,0"  ' "-200,500,-500,0,100,100,0"
        medInitLvwHead lvwBuildList, "작업번호,접수 번호", "-550,400"
        medInitLvwHead lvwAddList, "작업번호,접수 번호,환자 ID,성 명,검 체,검 사,비 고", _
                                   "-600,0,-400,-100,100,1000,-200"
    Else
        medInitLvwHead lvwPatient, "Patient ID,Patient Name,Sex/Age,Date of Birth,Location,Physician,Specimen", _
                                    "-100,300,-400,0,100,100,0"
        medInitLvwHead lvwPatient2, "Patient ID,Patient Name,Sex/Age,Date of Birth,Location,Physician,Specimen", _
                                    "-100,300,-400,0,100,100,0"
                                    
        medInitLvwHead lvwBuildList, "Work No,Accession No", "-550,550"
        medInitLvwHead lvwAddList, "Work No,Accession No,Patient ID,Patient Name,Specimen,Test,Remark", _
                                    "-300,-100,-600,-100,100,500,50"
        medInitLvwHead lvwLabNum, "Work No,Accession No,Patient ID,Patient Name,Specimen,Test,Remark", _
                            "-300,-100,-600,-100,100,500,50"
    
    End If
    '
End Sub

Private Sub ClearData()
    
    On Error Resume Next
    
    gstrMsk = "__-______-_____"
    
    txtWorkCd.Text = ""
    txtWorkCd.Enabled = True
    mskAccNo.Text = gstrMsk
    dtpWorkDt.Value = Date
    dtpWorkTm.Value = Format(Time, "hh:mm:ss")
    gstrPtAddInfo = ""
    '
    txtWorkCd.BackColor = vbWhite
    dtpWorkDt.Enabled = True
    dtpWorkTm.Enabled = True
    cmdBuild.Enabled = True
    cmdWSList.Enabled = True
    cmdExit.Enabled = True
    lblModifyLast.Enabled = True
    cmdSave.Enabled = False
    cmdAdd.Enabled = False
    If ActiveControl.Name <> optCondition(0).Name Then optCondition(0).Value = True
    optStatFg(2).Value = True
    '
    lvwBuildList.ListItems.Clear
    lvwAddList.ListItems.Clear
    lvwPatient.ListItems.Clear
    
    mskAccNo.BackColor = DCM_LightGray
    lvwBuildList.BackColor = DCM_LightGray
    lvwAddList.BackColor = DCM_LightGray
    lvwPatient.BackColor = DCM_LightGray
    
    fraWorkInfo.Enabled = False
    fraBuild.Enabled = False
    fraAdd.Enabled = False
    '
    lblCntBuild.Caption = ""
    lblCntAdd.Caption = ""
    lblCntTotal.Caption = ""
    lblFrSeq.Caption = ""
    lblToSeq.Caption = ""
    lblWorkCdNm.Caption = ""
    lblLastDtTm.Caption = ""

    
    lblSpcNo.Visible = False
    mskSpcNo.Visible = False
    lblLabNo.Visible = False
    lblLabNo.Caption = ""
    mskSpcNo.Text = "____________"
    chkSpcNo.Value = 0
    
    Set objLab301 = New clsWSBuild
    Set objDeleteList = New clsDictionary
    
    objDeleteList.Clear
    objDeleteList.DeleteAll
    objDeleteList.FieldInialize "workseq", "labno"
    '
End Sub

Private Sub EditData()
    '
    gstrPtAddInfo = ""
    mskAccNo.Text = gstrMsk
    txtWorkCd.BackColor = DCM_LightGray
    txtWorkCd.Enabled = False
    dtpWorkDt.Enabled = False
    dtpWorkTm.Enabled = False
    cmdBuild.Enabled = False
    cmdSave.Enabled = True
    cmdAdd.Enabled = False
    cmdWSList.Enabled = False
    lblModifyLast.Enabled = False
    '
    'fraWSHeader.Enabled = False
    fraWorkInfo.Enabled = True
    fraBuild.Enabled = True
    fraAdd.Enabled = True
    '
    mskAccNo.BackColor = vbWhite
    lvwBuildList.BackColor = vbWhite
    lvwAddList.BackColor = vbWhite
    lvwPatient.BackColor = vbWhite
    '
End Sub

Private Sub mskAccNo_KeyPress(KeyAscii As Integer)
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then mskAccNo_Validate (False)  'SendKeys "{TAB}"
End Sub

Private Sub mskAccNo_Validate(Cancel As Boolean)
    
    Dim strAddInfo  As String
    Dim strAccNo    As String
    Dim strDbAccNo  As String
    Dim ii          As Integer
    
    Dim Sess        As Object
    '
    cmdAdd.Enabled = False
    lvwPatient.ListItems.Clear
    gstrPtAddInfo = ""
    '
    If Trim(mskAccNo.Text) = gstrMsk Then Exit Sub
    '
    strDbAccNo = DBAccNo(mskAccNo.FormattedText)
    strAddInfo = objLab301.GetAddInfo(Sess, DBConn, strDbAccNo, _
                                      txtWorkCd.Text, DateStr(dtpWorkDt.Value), _
                                      Format(dtpWorkTm.Value, "hhmmss"))

    If strAddInfo = "" Then
        MsgBox "해당 자료가 없거나 부적절한 접수번호입니다 !", vbCritical
        mskAccNo.Text = gstrMsk
        Cancel = True
        mskAccNo.SetFocus
        Exit Sub
    Else
        'Duplication Check
        With lvwBuildList
            If .ListItems.Count > 0 Then
                For ii = 1 To .ListItems.Count
                    .ListItems(ii).Selected = True
                    If .ListItems(ii).ListSubItems.Item(1) = AccTrim(mskAccNo.FormattedText) Then
                        MsgBox "자료 중복 입력 Error!", vbInformation
                        mskAccNo.Text = gstrMsk
                        Cancel = True
                        mskAccNo.SetFocus
                        Exit Sub
                    End If
                Next
            End If
        End With
        With lvwAddList
            If .ListItems.Count > 0 Then
                For ii = 1 To .ListItems.Count
                    If .ListItems(ii).Text = AccTrim(mskAccNo.FormattedText) Then
                        MsgBox "자료 중복 입력 Error!", vbInformation
                        mskAccNo.Text = gstrMsk
                        Cancel = True
                        mskAccNo.SetFocus
                        Exit Sub
                    End If
                Next
            End If
        End With
        'Patient ListView Display
        gstrPtAddInfo = strAddInfo
        'Set objPtInfo = New clsPatientInfo
        With objPtInfo
            .PtType = RESULT_BY_DEFAULT
            .AccNo = AccTrim(mskAccNo.FormattedText)
            .LoadTable , ObjMyUser.EmpId
            If .RecordCount > 0 Then
                medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
                cmdAdd.Enabled = True
            Else
                mskAccNo.Text = gstrMsk
            End If
        End With
        'Set objPtInfo = Nothing
    End If
    DoEvents
    mskAddSeq.SetFocus

End Sub

Private Sub txtDatePicker_LostFocus()
    Call txtDatePicker_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub txtWorkCd_Change()
    If txtWorkCd.Text = "" Then lblWorkCdNm.Caption = ""
End Sub

Private Sub txtWorkCd_GotFocus()
    FocusMe Me.txtWorkCd
End Sub

Private Sub txtWorkCd_KeyPress(KeyAscii As Integer)
    
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then
        Call lstWSCode_KeyDown(vbKeyReturn, 0)
        lstWSCode.Visible = False
        Exit Sub
    ElseIf KeyAscii = vbKeyEscape Then
        lstWSCode.Visible = False
        Exit Sub
    End If

    If lstWSCode.ListCount > 0 Then
        lstWSCode.Visible = True
        lstWSCode.ZOrder 0
        Call medCodeHelp(KeyAscii, lstWSCode, txtWorkCd.Text, txtWorkCd, dtpWorkDt)
    End If

End Sub

Private Sub txtWorkCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstWSCode.ListCount = 0 Then Exit Sub
    Select Case KeyCode
    Case vbKeyDown:
        lstWSCode.Visible = True
        lstWSCode.ListIndex = 0
        lstWSCode.ZOrder 0
        lstWSCode.SetFocus
    Case vbKeyEscape:
        lstWSCode.Visible = False
    End Select
End Sub



Private Sub txtWorkCd_Validate(Cancel As Boolean)
    If txtWorkCd.Text = "" Then Exit Sub

    lblWorkCdNm.Caption = ""
    If Trim(txtWorkCd.Text) = "" Then
        Cancel = True
        txtWorkCd.SetFocus
        Exit Sub
    End If
    '
    If objLab301.IsWorkCd(txtWorkCd.Text) = False Then
        MsgBox "코드 입력 Error!", vbCritical
        Cancel = True
        txtWorkCd.SetFocus
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
            gstrLastDt = .Field3
            gstrLastTm = .Field4
            If gstrLastDt = "" Then    '10시간 전...
                gstrLastDt = Format(DateAdd("h", -10, Now), CS_DateDbFormat)
                gstrLastTm = Format(DateAdd("h", -10, Now), CS_TimeDbFormat)
            End If
            lblLastDtTm.Caption = Format(gstrLastDt, CS_DateMask) & "  " & Format(Mid(gstrLastTm, 1, 4), CS_TimeShortMask)
        End If
    End With
    Set objLab032 = Nothing
    lstWSCode.Visible = False
   '
End Sub

Private Sub cmdAdd_Click()
    
    Dim strTmp As String
    
    If lvwPatient.ListItems.Count = 0 Then Exit Sub
'    If mskAddSeq.ClipText = "" Then
'        MsgBox "추가검체의 Work No.를 입력하세요.", vbInformation, "메세지"
'        mskAddSeq.SetFocus
'        Exit Sub
'    End If
    lvwPatient.ListItems(1).Selected = True
    With lvwPatient.ListItems(1)
        strTmp = mskAddSeq.ClipText & vbTab & AccTrim(mskAccNo.FormattedText) & vbTab & .Text & vbTab & _
                        .ListSubItems.Item(1) & vbTab & gstrPtAddInfo '& vbTab
    End With
    medDataLoadLvw lvwAddList, vbNewLine, vbTab, strTmp
    cmdAdd.Enabled = False
    lvwPatient.ListItems.Clear
    mskAccNo.Text = gstrMsk
    mskAccNo.SetFocus
    DisplayCount
    
    blnChange = True

End Sub

Private Sub DisplayCount()
    lblCntBuild.Caption = lvwBuildList.ListItems.Count
    lblCntAdd.Caption = lvwAddList.ListItems.Count
    lblCntTotal.Caption = Val(lblCntBuild.Caption) + Val(lblCntAdd.Caption)
    With lvwBuildList
        If .ListItems.Count > 0 Then
            .ListItems(1).Selected = True
            lblFrSeq.Caption = .ListItems(1).Text
            lblToSeq.Caption = .ListItems(.ListItems.Count).Text
        Else
            lblFrSeq.Caption = ""
            lblToSeq.Caption = ""
        End If
    End With
End Sub

Private Function DBAccNo(ByVal strval As String) As String
    
    Dim aryTmp()    As String
    Dim ii          As Integer
    Dim intLen      As Integer
    
    aryTmp = Split(strval, "-")
    For ii = 1 To 2
        aryTmp(ii) = Trim(aryTmp(ii))
    Next
    If Mid(aryTmp(1), 1, 1) = "9" Then
        aryTmp(1) = "19" & aryTmp(1)
    Else
        aryTmp(1) = "20" & aryTmp(1)
    End If
   '
    DBAccNo = Join(aryTmp, "-")
    
End Function

Private Function AccTrim(ByVal strval As String) As String
    
    Dim aryTmp()    As String
    Dim ii          As Integer
    
    aryTmp = Split(strval, "-")
    For ii = 0 To 2
        aryTmp(ii) = Trim(aryTmp(ii))
    Next
    AccTrim = Join(aryTmp, "-")
    
End Function
Private Sub mskSpcNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskSpcNo_LostFocus
End Sub

Private Sub mskSpcNo_LostFocus()

    Dim sSpcYy      As String
    Dim sSpcNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String


    If Len(mskSpcNo.ClipText) > 0 And Len(mskSpcNo.ClipText) < 12 Then
        mskSpcNo.SetFocus
    End If
    If Trim(mskSpcNo.ClipText) = "" Then Exit Sub
    If txtWorkCd.Text = "" Then
        MsgBox "WorkSheet Code를 입력하세요", vbInformation + vbOKOnly
        mskSpcNo.Text = "____________"
        If txtWorkCd.Enabled Then txtWorkCd.SetFocus
        Exit Sub
    End If
    
    If Trim(mskSpcNo.ClipText) = "" Then Exit Sub
    
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    
    sSpcYy = Mid(mskSpcNo.ClipText, 1, P_SpcYyLength)
    sSpcNo = Format(Mid(mskSpcNo.ClipText, P_SpcYyLength + 1, P_SpcNoLength), "#0")
    
    lblLabNo.Caption = objPtInfo.GetLabNoBySpcNo(sSpcYy, sSpcNo)

    sWorkArea = medGetP(lblLabNo.Caption, 1, "-")
    sAccDt = medGetP(lblLabNo.Caption, 2, "-")
    If Len(sAccDt) = 8 Then sAccDt = Mid(sAccDt, 3)

    sAccSeq = medGetP(lblLabNo.Caption, 3, "-")
    lblLabNo.Caption = sWorkArea & "-" & sAccDt & "-" & sAccSeq
     
    Dim RS   As Recordset
    Dim WSeq As Long
    Dim SSQL As String
    
    SSQL = " SELECT a.workarea,a.accdt,a.accseq " & _
           " FROM " & T_LAB008 & " c," & T_LAB302 & " b," & T_LAB201 & " a" & _
           " WHERE " & _
            DBW("a.spcyy=", sSpcYy) & " AND " & DBW("a.spcno=", sSpcNo) & _
           " AND  a.stscd in('2','3')" & _
           " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
           " AND " & DBW("c.workcd=", txtWorkCd.Text) & " AND b.testcd=c.testcd AND b.spccd=c.spccd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        If lvwBuildList.ListItems.Count < 1 Then
            If objLab301 Is Nothing Then
                Set objLab301 = New clsWSBuild
            Else
                Set objLab301 = Nothing
                Set objLab301 = New clsWSBuild
            End If
            WSeq = objLab301.StartWorkSeq(txtWorkCd.Text, Format(GetSystemDate, "YYYYMMDD"))
        Else
            WSeq = lvwBuildList.ListItems(lvwBuildList.ListItems.Count).Text + 1
        End If
        
        Dim itmFound    As ListItem
        Dim iTmx        As ListItem
        
        Set itmFound = lvwBuildList.FindItem(lblLabNo.Caption, lvwSubItem)
        
        If itmFound Is Nothing Then
            Set iTmx = lvwBuildList.ListItems.Add
            iTmx.Text = WSeq
            iTmx.SubItems(1) = lblLabNo.Caption
            mskSpcNo.Text = "____________": If mskSpcNo.Enabled Then mskSpcNo.SetFocus
        Else
            MsgBox "이미 작성 대기중입니다.", vbInformation + vbOKOnly, "업무나열서"
            mskSpcNo.Text = "____________": If mskSpcNo.Enabled Then mskSpcNo.SetFocus
        End If
        Set itmFound = Nothing
    Else
        MsgBox "조건에 맞는 자료가 없습니다.", vbInformation + vbOKOnly, "업무나열서"
        mskSpcNo.Text = "____________": If mskSpcNo.Enabled Then mskSpcNo.SetFocus
    End If
    Set RS = Nothing
End Sub

' [이하부분은 추가 사항입니다.]
' 작성자 : KJG
' 내용   : 접수번호별로 세부사항 조회기능

Private Sub cmdSort_Click()
    Dim iTmx        As ListItem
    Dim strTmp      As String
    
    Dim lngStart    As Integer
    Dim lngEnd      As Integer
    Dim ii          As Integer
    
    optDiv(1).Value = True
    lngStart = lvwBuildList.ListItems.Item(1).Text
    lngEnd = lvwBuildList.ListItems.Item(lvwBuildList.ListItems.Count).Text
    
    Call lvwBuildList_Click
    
    With lvwLabNum
        .SortKey = 6
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    lvwBuildList.ListItems.Clear
    For ii = 1 To lvwLabNum.ListItems.Count
        Set iTmx = lvwBuildList.ListItems.Add(, , lngStart)
        iTmx.SubItems(1) = lvwLabNum.ListItems(ii).SubItems(1)
        lngStart = lngStart + 1
    Next
    
    
    Me.MousePointer = 11
    If optCondition(1).Value Then
        Dim objWSSave       As clsLisSqlResult
        Dim objProBar       As clsProgress
        Dim Resp            As VbMsgBoxResult
        Dim blnSave         As Boolean
        Dim strBuildList    As String
        Dim lngStartWSSeq   As Long
        Resp = MsgBox("WorkSheet를 재작성 하시겠습니까?", vbInformation + vbYesNo, "Info")
        If Resp = vbYes Then
            Set objProBar = New clsProgress
            With objProBar
                .Container = Me
                .Left = fraAdd.Left
                .Top = fraAdd.Top
                .Width = fraAdd.Width  ' (tabView.Width - 1700)
                .Height = 260
                .Message = "WORKSHEET 내역을 저장 중입니다..."
                .Max = 90
                
'                .SetMyForm Me
'                .Choice = True
'                .XPos = fraAdd.Left
'                .YPos = fraAdd.Top
'                .XWidth = fraAdd.Width  ' (tabView.Width - 1700)
'                .ForeColor = &H864B24
'                .Appearance = aPlate
'                .BorderStyle = bsNone
'                .YHeight = 260
'                .MSG = "Worksheet 내역을 저장 중입니다...."
'                .Max = 90
'                .Value = 1
                DoEvents
            End With
            Set objWSSave = New clsLisSqlResult
            
            blnSave = True
            
            For ii = 1 To lvwLabNum.ListItems.Count
                If ii = 1 Then
                    strTmp = lvwLabNum.ListItems.Item(ii).Text
                Else
                    strTmp = strTmp & COL_DIV & lvwLabNum.ListItems.Item(ii).Text
                End If
            Next
            
            If Not objWSSave.SetWorkSheetDelete(Trim(txtWorkCd.Text), Format(dtpWorkDt.Value, CS_DateDbFormat), strTmp) Then
                blnSave = False
            End If
            strBuildList = ""
            If blnSave Then
            
                For ii = 1 To lvwBuildList.ListItems.Count
                    If ii = 1 Then lngStartWSSeq = lvwBuildList.ListItems.Item(ii).Text
                    strTmp = DBAccNo(lvwBuildList.ListItems.Item(ii).SubItems(1)) & COL_DIV & lvwBuildList.ListItems.Item(ii).Text & COL_DIV & ""
                    If strBuildList <> "" Then
                        strBuildList = strBuildList & vbTab & strTmp
                    Else
                        strBuildList = strTmp
                    End If
                Next ii
            
                With objWSSave
                    If strBuildList <> "" Then
                        Call .BuildWorkSheet(Format(dtpWorkDt.Value, CS_DateDbFormat), txtWorkCd.Text, _
                                             Format(Time, "hhmmss"), lngStartWSSeq, ObjSysInfo.EmpId, strBuildList, _
                                             objLab301.GetLastDt, objLab301.GetLastTm, objProBar)
                        If Not .SaveOk = True Then
                            MsgBox "WorkSheet 작성도중 Error가 발생하였습니다.", vbInformation + vbOKOnly, "Info"
                        Else
                            Resp = MsgBox("지금 출력하시겠습니까 ? ", vbYesNo + vbInformation, "워크쉬트 출력")
                            If Resp = vbYes Then
                                objProBar.Message = "Worksheet List를 츨력 중입니다... 잠시만 기다리세요! "
                                DoEvents
                                
                                If Printers.Count = 0 Then
                                    MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
                                    Exit Sub
                                End If
                                
                                Dim MyReport As New clsWorkListG
                                With MyReport
                                    .WorkCode = txtWorkCd.Text
                                    .WorkName = lblWorkCdNm.Caption
                                    .WorkDate = Format(dtpWorkDt.Value, CS_DateDbFormat)
                                    .FromSeq = lblFrSeq.Caption
                                    .ToSeq = lblToSeq.Caption
                                    Call .Print_Worksheet
                                End With
                            
                                Set MyReport = Nothing
                            End If
                        End If
                    End If
                End With
                Call ClearData: Call cmdMin_Click
                Set objWSSave = Nothing
                Set objProBar = Nothing
            End If
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub lvwBuildList_Click()
    Dim objPrgBar As clsProgress
    
    Dim strLabNo    As String
    Dim strWorkNo   As String
    Dim ii          As Integer
    
    DoEvents
    If fraBuild.Width < 5000 Then Exit Sub
    
    Set objPrgBar = New clsProgress
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = medMain.stsBar
    objPrgBar.Max = lvwBuildList.ListItems.Count
    
    lvwPatient2.ListItems.Clear
    lvwLabNum.ListItems.Clear
    
    With lvwLabNum
        .SortKey = 6
        .SortOrder = lvwAscending
        .Sorted = False
    End With
    
    Me.MousePointer = 11
    If optDiv(0).Value Then
        strWorkNo = lvwBuildList.ListItems(lvwBuildList.SelectedItem.Index).Text
        strLabNo = lvwBuildList.ListItems(lvwBuildList.SelectedItem.Index).SubItems(1)
        Call GetPatientInfo(strWorkNo, strLabNo)
    Else
        For ii = 1 To lvwBuildList.ListItems.Count
            strWorkNo = lvwBuildList.ListItems(ii).Text
            strLabNo = lvwBuildList.ListItems(ii).SubItems(1)
            Call GetPatientInfo(strWorkNo, strLabNo)
            objPrgBar.Message = strLabNo & "의 내역을 조회중입니다."
            objPrgBar.Value = ii
        Next
    End If
        
    With objPtInfo
        .PtType = RESULT_BY_DEFAULT
        .AccNo = AccTrim(lvwLabNum.ListItems(1).SubItems(1))
        .LoadTable , ObjMyUser.EmpId
        If .RecordCount > 0 Then
            medDataLoadLvw lvwPatient2, vbNewLine, vbTab, .GetStringPtInfo
        End If
    End With
    Set objPrgBar = Nothing
    Me.MousePointer = 0
End Sub

Private Sub GetPatientInfo(ByVal strWorkNo As String, ByVal strLabNo As String)
    With objPtInfo
        .PtType = RESULT_BY_DEFAULT
        .AccNo = AccTrim(strLabNo)
        .LoadTable , ObjMyUser.EmpId
        If .RecordCount > 0 Then
            medDataLoadLvw lvwPatient2, vbNewLine, vbTab, .GetStringPtInfo
            Call GetDetailnfo(strWorkNo, strLabNo, medGetP(.GetStringPtInfo, 1, vbTab), medGetP(.GetStringPtInfo, 2, vbTab))
        End If
        lvwPatient2.ListItems.Clear
    End With
End Sub

Private Sub GetDetailnfo(ByVal WorkNo As String, ByVal strLabNo As String, ByVal sPtid As String, ByVal sPtNm As String)
    Dim strDbAccNo  As String
    Dim strTmp      As String
    Dim Sess        As Object
    '
    gstrPtAddInfo = ""
    '
    If strLabNo = "" Then Exit Sub
    
    strDbAccNo = DBAccNo(strLabNo)
    gstrPtAddInfo = objLab301.GetAddInfo(Sess, DBConn, strDbAccNo, _
                                      txtWorkCd.Text, DateStr(dtpWorkDt.Value), Format(dtpWorkTm.Value, "hhmmss"))
    strTmp = WorkNo & vbTab & strLabNo & vbTab & sPtid & vbTab & sPtNm & vbTab & gstrPtAddInfo '& vbTab
    medDataLoadLvw lvwLabNum, vbNewLine, vbTab, strTmp

End Sub



