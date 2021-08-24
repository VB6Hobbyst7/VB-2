VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmOnLineHelp 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사항목 세부설명"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ClipControls    =   0   'False
   Icon            =   "frmOnLineHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "확인(&O)"
      Height          =   495
      Left            =   6975
      Style           =   1  '그래픽
      TabIndex        =   34
      Top             =   7230
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   3
      Left            =   4350
      TabIndex        =   10
      Top             =   45
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 상세 정보1"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   3225
      Left            =   4350
      TabIndex        =   9
      Top             =   300
      Width           =   3945
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   8
         Left            =   30
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "소요시간(S)"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   9
         Left            =   30
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2775
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "소요시간(T)"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   6
         Left            =   30
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1650
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   7
         Left            =   30
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2025
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   4
         Left            =   30
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   900
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "처방일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   5
         Left            =   30
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   2
         Left            =   30
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "처방명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   3
         Left            =   30
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   525
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "검체명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrderNm 
         Height          =   360
         Left            =   1335
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   150
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   360
         Left            =   1350
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   525
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblOrdDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   900
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblColDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1275
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
         BackColor       =   14411494
         ForeColor       =   255
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
      Begin MedControls1.LisLabel lblRcvDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1650
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
         BackColor       =   14411494
         ForeColor       =   16744576
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
      Begin MedControls1.LisLabel lblVfyDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2025
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
         BackColor       =   14411494
         ForeColor       =   4210752
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
      Begin MedControls1.LisLabel lblTAT 
         Height          =   360
         Left            =   1350
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblTAT1 
         Height          =   360
         Left            =   1350
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2775
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   635
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
         Caption         =   ""
      End
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검사정보"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   930
      Left            =   30
      TabIndex        =   2
      Top             =   300
      Width           =   4305
      Begin MedControls1.LisLabel lblTestCd 
         Height          =   360
         Left            =   1395
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   375
         Left            =   1395
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   510
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   661
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   0
         Left            =   60
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   135
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   1
         Left            =   60
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Caption         =   "검사명"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   2
      Left            =   30
      TabIndex        =   1
      Top             =   1230
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 검체정보"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   3540
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 참고치 정보"
      LeftGab         =   100
   End
   Begin VB.Frame fraRefRange 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3420
      Left            =   30
      TabIndex        =   4
      Tag             =   "35208"
      Top             =   3780
      Width           =   4305
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   405
         Left            =   60
         TabIndex        =   5
         Top             =   450
         Width           =   4080
         Begin MSComctlLib.TabStrip tabRefAppDt 
            Height          =   300
            Left            =   60
            TabIndex        =   6
            Top             =   75
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   529
            Style           =   2
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
         Begin VB.Shape Shape2 
            BorderColor     =   &H0085A3A3&
            BorderWidth     =   2
            Height          =   345
            Left            =   45
            Shape           =   4  '둥근 사각형
            Top             =   60
            Width           =   4020
         End
      End
      Begin MSComctlLib.ListView lvwReference 
         Height          =   2070
         Left            =   75
         TabIndex        =   7
         Top             =   870
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3651
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15728382
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Range"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2730
         Top             =   150
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOnLineHelp.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOnLineHelp.frx":0144
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOnLineHelp.frx":027C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblItemSpec 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Ca - Whole Blood"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   4005
      End
   End
   Begin MedControls1.LisLabel lblAction 
      Height          =   330
      Index           =   4
      Left            =   4350
      TabIndex        =   11
      Top             =   3540
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   582
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 상세 정보2"
      LeftGab         =   100
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   3405
      Left            =   4335
      TabIndex        =   12
      Top             =   3795
      Width           =   3945
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   10
         Left            =   30
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
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
         Caption         =   "검사부서"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   11
         Left            =   30
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   495
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
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
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   12
         Left            =   30
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   870
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
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
         Caption         =   "검 사 일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Index           =   13
         Left            =   30
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
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
         Caption         =   "특이사항"
         Appearance      =   0
      End
      Begin VB.TextBox txtRemark 
         Height          =   1470
         Left            =   30
         TabIndex        =   33
         Top             =   1860
         Width           =   3870
      End
      Begin MedControls1.LisLabel lblSection 
         Height          =   375
         Left            =   1215
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   661
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblTel 
         Height          =   360
         Left            =   1215
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   495
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   635
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
         Caption         =   ""
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DBE6E6&
         Height          =   840
         Left            =   1215
         TabIndex        =   25
         Top             =   780
         Width           =   2685
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "일요일"
            Height          =   180
            Index           =   6
            Left            =   75
            TabIndex        =   32
            Top             =   570
            Width           =   840
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "토요일"
            Height          =   180
            Index           =   5
            Left            =   1815
            TabIndex        =   31
            Top             =   360
            Width           =   840
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "금요일"
            Height          =   180
            Index           =   4
            Left            =   930
            TabIndex        =   30
            Top             =   360
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "목요일"
            Height          =   180
            Index           =   3
            Left            =   75
            TabIndex        =   29
            Top             =   360
            Width           =   840
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "수요일"
            Height          =   180
            Index           =   2
            Left            =   1815
            TabIndex        =   28
            Top             =   150
            Width           =   840
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "화요일"
            Height          =   180
            Index           =   1
            Left            =   930
            TabIndex        =   27
            Top             =   150
            Width           =   870
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00DBE6E6&
            Caption         =   "월요일"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   26
            Top             =   150
            Width           =   840
         End
      End
   End
   Begin FPSpread.vaSpread tblSpcList 
      Height          =   1935
      Left            =   30
      TabIndex        =   35
      Tag             =   "35220"
      Top             =   1575
      Width           =   4275
      _Version        =   196608
      _ExtentX        =   7541
      _ExtentY        =   3413
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
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
      MaxCols         =   5
      MaxRows         =   8
      OperationMode   =   1
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmOnLineHelp.frx":03B4
      VirtualRows     =   7
   End
End
Attribute VB_Name = "frmOnLineHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarTestCd   As String
Private mvarTestNm   As String
Private mvarSpccd    As String
Private mvarWorkarea As String
Private mvarAccdt    As String
Private mvarAccSeq   As String
Private mvarVfyDate  As String
Private mvarRcvDate  As String
Private mvarOrdDate  As String
Private mvarTAT      As String


Public Property Let TestCd(ByVal vData As String)
    mvarTestCd = vData
End Property
Public Property Let TestNm(ByVal vData As String)
    mvarTestNm = vData
End Property
Public Property Let SpcCd(ByVal vData As String)
    mvarSpccd = vData
End Property

Public Property Let AccDt(ByVal vData As String)
    mvarAccdt = vData
End Property

Public Property Let WorkArea(ByVal vData As String)
    mvarWorkarea = vData
End Property

Public Property Let AccSeq(ByVal vData As String)
    mvarAccSeq = vData
End Property
Public Property Let VfyDate(ByVal vData As String)
    mvarVfyDate = vData
End Property
Public Property Let RcvDate(ByVal vData As String)
    mvarRcvDate = vData
End Property

Public Property Let OrdDate(ByVal vData As String)
    mvarOrdDate = vData
End Property
Public Property Let TAT(ByVal vData As String)
    mvarTAT = vData
End Property
Private Sub ClearData(Optional ByVal blnFirst As Boolean = False)
    Dim I As Integer
    
    lblTestCd.Caption = "": lblOrderNm.Caption = ""
    lblTestNm.Caption = "": lblRcvDate.Caption = ""
    lblVfyDate.Caption = "": lblOrdDate.Caption = ""
    lblTAT.Caption = "":    lblSection.Caption = "": lblColDate.Caption = ""
    lblTel.Caption = "": lblTAT1.Caption = ""
    txtRemark.Text = ""
    
    For I = 0 To 6
        chkDay(I).Value = 0
    Next
    If blnFirst = True Then
        Call medClearTable(tblSpcList)
        tabRefAppDt.Tabs.Clear
        lvwReference.ListItems.Clear
    End If
    
End Sub


Private Sub Query()
    If mvarTestCd = "" Then Exit Sub
    lblTestCd.Caption = mvarTestCd
    lblTestNm.Caption = mvarTestNm: lblOrderNm.Caption = mvarTestNm
    lblRcvDate.Caption = mvarRcvDate
    lblVfyDate.Caption = mvarVfyDate
    lblOrdDate.Caption = mvarOrdDate
    lblTAT.Caption = mvarTAT
    
    '검사코드에 대한 지정검체를 조회한다
    Call LabSpecimenLoad
    '채혈일자구하기
'    Call ShowColDate '(채혈일자구하기)
'    '지정검체에별 참고치를 가지고 온다
'     Call LoadRefData(mvarTestCd, mvarSpccd)
'    '검사항목별 세부정보 구하시
'    Call ShowTestInfo(mvarTestCd, mvarSpccd)
End Sub

Private Sub ShowTestInfo(ByVal sTestCd As String, ByVal sSpcCd As String)
    Dim SSQL        As String
    Dim RS          As Recordset
    Dim strAry()    As String
    Dim ii          As Integer
    
    SSQL = " select * from " & T_LAB031 & _
           " where " & DBW("cdindex=", LC4_TestItemComment) & _
           " and " & DBW("cdval1=", sTestCd) & _
           " and " & DBW("cdval2=", sSpcCd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblSection.Caption = RS.Fields("field3").Value & ""
        lblTel.Caption = RS.Fields("field5").Value & ""
        If RS.Fields("field4").Value & "" <> "" Then
            If RS.Fields("field4").Value & "" = "7" Then
                For ii = 0 To 6
                    chkDay(ii).Value = 1
                Next
            Else
                strAry() = Split(RS.Fields("field4").Value & "", COL_DIV)
                For ii = 0 To UBound(strAry)
                    chkDay(strAry(ii)).Value = 1
                Next
            End If
        End If
        txtRemark.Text = RS.Fields("text1").Value & ""
    End If
    Set RS = Nothing
End Sub

Private Sub ShowColDate()
    Dim SSQL As String
    Dim RS   As Recordset
    
On Error Resume Next
    SSQL = " select coldt,coltm from " & T_LAB201 & _
           " where " & DBW("workarea=", mvarWorkarea) & _
           " and " & DBW("accdt=", mvarAccdt) & _
           " and " & DBW("accseq=", mvarAccSeq)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblColDate.Caption = Format(RS.Fields("coldt").Value & "", "####-##-##") & " " & Format(Mid(RS.Fields("coltm").Value & "", 1, 4), "0#:##")
    End If
    Set RS = Nothing
        
End Sub

'****************************************************
'                   지정검체 가져오기
'****************************************************
Private Function SqlSpecimenRead() As String
    Dim SSQL As String
    
    SSQL = " select a.workarea,b.spccd,c.field4 as spcnm,d.field3 as workareanm,b.seq " & _
           " from " & T_LAB001 & " a," & T_LAB004 & " b," & T_LAB032 & " c," & T_LAB032 & " d" & _
           " where  " & DBW("a.testcd=", mvarTestCd) & _
           " and a.applydt=(select max(applydt) from s2lab001 where testcd=a.testcd)" & _
           " and (a.expdt='' or a.expdt is null)" & _
           " and a.testcd=b.testcd" & _
           " and " & DBW("c.cdindex=", LC3_Specimen) & _
           " and c.cdval1=b.spccd" & _
           " and b.applydt=(select max(applydt) from s2lab004 where testcd=b.testcd and spccd=b.spccd)" & _
           " and (b.expdt='' or b.expdt is null)" & _
           " and " & DBW("d.cdindex=", LC3_WorkArea) & _
           " and d.cdval1=a.workarea" & _
           " order by seq"
    SqlSpecimenRead = SSQL
End Function

'% Sub Routine 3 : LabSpecimenLoad
'%                 지정검체명들을 Tab에 Display

Private Sub LabSpecimenLoad()

    Dim RS      As Recordset
    Dim I           As Integer
    Dim strSpcnm    As String
    
    Set RS = New Recordset
    RS.Open SqlSpecimenRead, DBConn
    
    With tblSpcList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 0
    
        .Row = 0
        While (RS.EOF = False)
            If .Row = .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .Row + 1
            .Col = 2: .Value = Trim("" & RS.Fields("seq").Value)
            .Col = 4: .Value = Trim("" & RS.Fields("spcnm").Value): strSpcnm = .Value
            .Col = 3: .Value = Trim("" & RS.Fields("spccd").Value)
                      
            If mvarSpccd = .Value Then
                .Col = 1: .Value = "▶": .ForeColor = DCM_LightRed
                lblSpcNm.Caption = strSpcnm
                lblItemSpec.Caption = mvarTestNm & "-" & strSpcnm
                Call tblSpcList_Click(1, .Row)
            End If
            .Col = 5: .Value = Trim("" & RS.Fields("workarea").Value & Space(5) & RS.Fields("workareanm").Value & "")
            RS.MoveNext
        Wend
        .RowHeight(-1) = 13.3
    End With

    Set RS = Nothing

End Sub

Private Sub LoadRefData(ByVal sTestCd As String, ByVal sSpcCd As String)
    tabRefAppDt.Tabs.Clear
    lvwReference.ListItems.Clear
    
    Call DisplayRefApplyDt(sTestCd, sSpcCd)
    
    If tabRefAppDt.Tabs.Count > 0 Then
        Call tabRefAppDt_Click
    End If
End Sub
Private Function GetRefApplydtSQL(ByVal sTestCd As String, ByVal sSpcCd As String) As String
    GetRefApplydtSQL = " Select applydt " & _
                       " From  " & T_LAB005 & _
                       " Where " & DBW("testcd", sTestCd, 2) & _
                       " and   " & DBW("spccd", sSpcCd, 2) & _
                       " Group by applydt " & _
                       " Order by applydt"
End Function

Private Sub DisplayRefApplyDt(ByVal sTestCd As String, ByVal sSpcCd As String)

    Dim RS       As Recordset       'Oracle DynaSet
    Dim tmpKey      As String
    Dim tmpCaption  As String
    Dim tmpSQL      As String
    Dim I           As Integer

    tmpSQL = GetRefApplydtSQL(sTestCd, sSpcCd)
    
    Set RS = New Recordset
    RS.Open tmpSQL, DBConn

    I = 0
    tabRefAppDt.Tabs.Clear
    While (RS.EOF = False)
        I = I + 1
        tmpKey = "" & RS.Fields("ApplyDt").Value
        tmpCaption = Format(tmpKey, CS_DateMask)
        tabRefAppDt.Tabs.Add I, , tmpCaption
        RS.MoveNext
    Wend

    Set RS = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call ClearData(True)
    Call Query
End Sub

Private Sub tabRefAppDt_Click()

    Dim tmpSQL As String
    Dim tmpAppDt As String
    Dim tmpSpcCd As String
    
    tmpAppDt = Format(tabRefAppDt.SelectedItem.Caption, CS_DateDbFormat)
    tmpSpcCd = mvarSpccd
    tmpSQL = GetRefAllDataSQL(lblTestCd.Caption, tmpSpcCd, tmpAppDt)
    Call ShowRefAllData(tmpSQL)

End Sub

Private Function GetRefAllDataSQL(ByVal TestCd As String, ByVal SpcCd As String, _
                                 ByVal ApplyDt As String) As String
    GetRefAllDataSQL = " Select * " & _
                       " From  " & T_LAB005 & _
                       " Where " & DBW("testcd", TestCd, 2) & _
                       " and   " & DBW("spccd", SpcCd, 2) & _
                       " and   " & DBW("applydt", ApplyDt, 2) & _
                       " Order by applysex, agefrom"
End Function

Private Sub ShowRefAllData(ByVal SqlStmt As String)

    Dim intFieldCount As Integer
    Dim aryTitle As Variant
    Dim aryWidth As Variant
    Dim itmx As ListItem
    Dim RS  As Recordset
    Dim I As Integer
    
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.EOF Then GoTo NoData

    I = 0
    intFieldCount = 5
    aryTitle = Array("성별", "일령", "기준치", "Auto Value", "Panic Value")
    aryWidth = Array(-85, -55, 50, -50, 0)
    lvwReference.ColumnHeaders.Clear
    With lvwReference
        For I = 1 To intFieldCount
            .ColumnHeaders.Add I, aryTitle(I - 1), aryTitle(I - 1), (.Width \ intFieldCount) _
                                + aryWidth(I - 1), vbLeftJustify
        Next I
        .View = lvwReport                          ' lvwReport = 3 (Report Style)
    End With
    lvwReference.ListItems.Clear
    lvwReference.SmallIcons = ImageList1
    lvwReference.Icons = ImageList1

    While (RS.EOF = False)
        I = I + 1
        Select Case "" & RS.Fields("ApplySex").Value
        Case "M":
            Set itmx = lvwReference.ListItems.Add(, , "남자", 1, 1)
        Case "F":
            Set itmx = lvwReference.ListItems.Add(, , "여자", 2, 2)
        Case "B":
            Set itmx = lvwReference.ListItems.Add(, , "Both")
        Case "U":
            Set itmx = lvwReference.ListItems.Add(, , "Unknown", 3)
        End Select
        itmx.SubItems(1) = "" & RS.Fields("AgeFrom").Value & " - " & RS.Fields("AgeTo").Value & " days"
        '선린 : 2001-05-31(추가)
        If Val("" & RS.Fields("RefValFrom").Value) = 0 And Val("" & RS.Fields("RefValTo").Value) = 0 Then
            itmx.SubItems(2) = "" & RS.Fields("RefCd").Value
        Else
            itmx.SubItems(2) = "" & RS.Fields("RefValFrom").Value & " - " & RS.Fields("RefValTo").Value
            If Len("" & RS.Fields("RefCd").Value) Then
                itmx.SubItems(2) = itmx.SubItems(2) & "(" & RS.Fields("RefCd").Value & ")"
            End If
        End If
        itmx.SubItems(3) = "" & RS.Fields("aRefValFrom").Value & " - " & RS.Fields("aRefValTo").Value
        itmx.SubItems(4) = "" & RS.Fields("panicfrval").Value & " - " & RS.Fields("panictoval").Value
        RS.MoveNext
    Wend


NoData:
    Set RS = Nothing
    Exit Sub

End Sub

Private Sub tblSpcList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii          As Integer
    Dim strTestcd   As String
    Dim strSpcCd    As String
    
    Dim strTmp      As String
    Dim strMin      As String
    Dim strHour     As String
    Dim strDay      As String
    
    If Row < 1 Then Exit Sub
    Call ClearData
    lblTestCd.Caption = mvarTestCd
    lblTestNm.Caption = mvarTestNm: lblOrderNm.Caption = mvarTestNm
    With tblSpcList
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = 1
        .BlockMode = True
        .Text = ""
        .BlockMode = False
        .Row = Row: .Col = 1
        .Value = "▶": .ForeColor = DCM_LightRed
        strTestcd = mvarTestCd
        .Col = 3: strSpcCd = .Value
        .Col = 4: lblSpcNm.Caption = .Value
        
    End With
    '채혈일자구하기
    If strSpcCd = mvarSpccd Then
        Call ShowColDate
        lblRcvDate.Caption = mvarRcvDate
        lblVfyDate.Caption = mvarVfyDate
        lblOrdDate.Caption = mvarOrdDate
        lblTAT.Caption = mvarTAT
        
        If lblRcvDate.Caption <> "" Then
            If lblVfyDate.Caption <> "" Then
                strTmp = DateDiff("n", lblRcvDate.Caption, lblVfyDate.Caption)
            Else
                strTmp = DateDiff("n", lblRcvDate.Caption, GetSystemDate)
            End If
            If strTmp >= 60 * 24 Then
                strDay = strTmp \ (60 * 24)
                strHour = (strTmp Mod (60 * 24)) Mod 24
            ElseIf strTmp >= 60 Then
                strHour = strTmp \ 60
            End If
            strMin = strTmp Mod 60
            
            If strDay <> "" Then
                lblTAT1.Caption = "접수 후 " & strDay & "일 " & IIf(strHour = "", "0", strHour) & "시간 " & strMin & "분 경과됨."
            Else
                If strHour <> "" Then
                    lblTAT1.Caption = "접수 후 " & IIf(strHour = "", "0", strHour) & "시간 " & strMin & "분 경과됨."
                Else
                    lblTAT1.Caption = "접수 후 " & strMin & "분 경과됨."
                End If
            End If
            
        End If
        
    End If
    '지정검체에별 참고치를 가지고 온다
     Call LoadRefData(strTestcd, strSpcCd)
    '검사항목별 세부정보 구하기
    Call ShowTestInfo(strTestcd, strSpcCd)
End Sub
