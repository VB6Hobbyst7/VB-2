VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form medTelephone 
   BackColor       =   &H00DBE6E6&
   Caption         =   "병실 전화번호"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "돋움체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "medTelephone.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   9945
   StartUpPosition =   1  '소유자 가운데
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00DBE6E6&
      Caption         =   "조회"
      Height          =   435
      Index           =   1
      Left            =   90
      Style           =   1  '그래픽
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1005
   End
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00DBE6E6&
      Caption         =   "등록"
      Height          =   435
      Index           =   0
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8475
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   6090
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabTelNo 
      Height          =   345
      Left            =   1155
      TabIndex        =   26
      Top             =   2190
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "진료과"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "병동Station"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "4층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "5층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "6층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "7층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "8층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "9층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "10층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "11층"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "12층"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread tblward 
      Height          =   3420
      Left            =   1140
      TabIndex        =   27
      Top             =   2550
      Width           =   5940
      _Version        =   196608
      _ExtentX        =   10478
      _ExtentY        =   6033
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   10
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "medTelephone.frx":09EA
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   10
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   7140
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2220
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   556
      BackColor       =   12713468
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
      Caption         =   "진료과 전화번호"
   End
   Begin FPSpread.vaSpread tbldept 
      Height          =   3420
      Left            =   7155
      TabIndex        =   29
      Top             =   2565
      Width           =   2655
      _Version        =   196608
      _ExtentX        =   4683
      _ExtentY        =   6033
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   10
      OperationMode   =   2
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "medTelephone.frx":0E6B
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   10
   End
   Begin VB.Frame fraSave 
      BackColor       =   &H00DBE6E6&
      Height          =   1380
      Left            =   1170
      TabIndex        =   4
      Top             =   765
      Width           =   8595
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장(&&S)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7185
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   780
         Width           =   1320
      End
      Begin VB.TextBox txtRoom 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6660
         TabIndex        =   11
         Top             =   330
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2175
         TabIndex        =   10
         Top             =   720
         Width           =   1245
      End
      Begin VB.OptionButton optBussdiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   315
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optBussdiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         Caption         =   "진료과"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   75
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   780
         Width           =   945
      End
      Begin VB.CommandButton cmdPopupList 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3420
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "medTelephone.frx":1271
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   315
         Width           =   300
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   5
         Top             =   315
         Width           =   1245
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   360
         Left            =   3735
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   330
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   1125
         TabIndex        =   40
         Top             =   315
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "병동코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   1125
         TabIndex        =   41
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "전화번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRoom 
         Height          =   360
         Left            =   5625
         TabIndex        =   42
         Top             =   330
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "병상정보"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel lblcode 
      Height          =   360
      Left            =   90
      TabIndex        =   30
      Top             =   2190
      Width           =   1005
      _ExtentX        =   1773
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
      Caption         =   "정보조회"
      Appearance      =   0
   End
   Begin VB.Frame fraReview 
      BackColor       =   &H00DBE6E6&
      Height          =   1380
      Left            =   1170
      TabIndex        =   13
      Top             =   765
      Width           =   8595
      Begin VB.TextBox txtPtid 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   14
         Top             =   180
         Width           =   1185
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   2265
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lblsexage 
         Height          =   345
         Left            =   4665
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lbldob 
         Height          =   345
         Left            =   7020
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lbllocation 
         Height          =   345
         Left            =   7020
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lblbedindt 
         Height          =   345
         Left            =   4665
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   555
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lblTestnm 
         Height          =   360
         Left            =   1080
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   930
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   635
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lblphone 
         Height          =   360
         Left            =   7020
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   930
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         BackColor       =   14351358
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
      End
      Begin MedControls1.LisLabel lblorddt 
         Height          =   345
         Left            =   4665
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   930
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         BackColor       =   14351358
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
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1080
         TabIndex        =   23
         Top             =   465
         Width           =   2535
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "병동환자"
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
            Left            =   1290
            TabIndex        =   25
            Tag             =   "10109"
            Top             =   150
            Width           =   1140
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외래환자"
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
            Left            =   90
            TabIndex        =   24
            Tag             =   "10108"
            Top             =   150
            Width           =   1140
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   1
         Left            =   45
         TabIndex        =   31
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   2
         Left            =   45
         TabIndex        =   32
         Top             =   555
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   3
         Left            =   45
         TabIndex        =   33
         Top             =   930
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "최근처방"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   4
         Left            =   3630
         TabIndex        =   34
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Height          =   345
         Index           =   5
         Left            =   3630
         TabIndex        =   35
         Top             =   555
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "입원일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   6
         Left            =   3630
         TabIndex        =   36
         Top             =   930
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "처방일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   7
         Left            =   5970
         TabIndex        =   37
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Height          =   345
         Index           =   8
         Left            =   5970
         TabIndex        =   38
         Top             =   555
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "Location"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   9
         Left            =   5970
         TabIndex        =   39
         Top             =   930
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         Caption         =   "전화번호"
         Appearance      =   0
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "전화번호"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   900
      TabIndex        =   0
      Top             =   195
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   345
      Picture         =   "medTelephone.frx":17FB
      Top             =   135
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   600
      Index           =   0
      Left            =   135
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   2670
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '단색
      Height          =   600
      Index           =   1
      Left            =   180
      Shape           =   4  '둥근 사각형
      Top             =   90
      Width           =   2655
   End
End
Attribute VB_Name = "medTelephone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents objCodeList  As clspopuplist
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call TabIni
End Sub

Private Sub Form_Load()
    optDiv(0).Value = True
    fraSave.Visible = True: fraReview.Visible = False
    Call FormClear
End Sub

Private Sub TabIni()
    Dim iCnt As Long
    Dim SSQL As String
    Dim RS   As Recordset
'    Dim objData As clsBasisData
    
    Call medClearTable(tblward)
    Call medClearTable(tbldept)
    
'    Set objData = New clsBasisData
    Set RS = New Recordset
    RS.Open GetSQLWardList, DBConn
    
    tabTelNo.Tabs.Clear
'    iCnt = objLisComCode.WardId.RecordCount
    
    If RS.EOF Then Exit Sub
    
'    If iCnt < 1 Then Exit Sub
'    With objLisComCode.WardId
    With RS
        Do Until .EOF
            tabTelNo.Tabs.Add , .Fields("wardid").Value & "", .Fields("wardnm").Value & ""
            .MoveNext
        Loop
    End With

    tabTelNo.Tabs.Item(1).Selected = True
    
    SSQL = "SELECT field1,field2 FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC2_TelePhone) & " AND " & DBW("field3=", "1")
    
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    With tbldept
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = RS.Fields("field2").Value & ""
                          .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                .Col = 2: .Value = RS.Fields("field1").Value & ""
                          .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                RS.MoveNext
            Loop
        End If
    End With

    Set RS = Nothing
'    Set objData = Nothing
End Sub


Private Sub FormClear()
    Frame5.Enabled = False
    lblcode.Caption = ""
    txtCd.Text = ""
    txtRoom.Text = ""
    lblName.Caption = ""
    txtPtid.Text = "": lblPtNm.Caption = "": lblsexage.Caption = "": lblbedindt.Caption = ""
    lbllocation.Caption = "": lblTestnm.Caption = "": lblphone.Caption = "": lblorddt.Caption = ""
    txtPhone.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set objCodeList = Nothing
End Sub

Private Sub optBussdiv_Click(Index As Integer)
    Call FormClear
    If Index = 0 Then
        lblcode.Caption = "병동코드"
        lblRoom.Visible = True
        txtRoom.Visible = True
    Else
        lblcode.Caption = "진료과코드"
        lblRoom.Visible = False
        txtRoom.Visible = False
    End If
    txtCd.SetFocus
End Sub

Private Sub cmdPopupList_Click()
    Dim lngTop  As Long
    Dim lngLeft As Long
    Dim tmpSQL  As String

'    Set objCodeList = New clspopuplist
    Set objCodeList = New clsPopUpList
    
    With objCodeList
        .Connection = DBConn
        .FormCaption = "리스트"
        .ColumnHeaderText = "코드;코드명"
        
        If optBussdiv(0).Value Then
           .LoadPopUp GetSQLWardList
        Else
           .LoadPopUp GetSQLDeptList
        End If
        
        txtCd.Text = .SelectedItems(0)
        lblName.Caption = .SelectedItems(1)
        
'        lngTop = txtCd.Top + fraSave.Top
'        lngLeft = Me.Left + txtCd.Left + fraSave.Left
'        .Caption = "리스트"
'        .HeadName = "코드,코드명"
        
'        If optBussdiv(0).Value Then
'           .ListPop GETSQLWARD, lngTop, lngLeft ', objLisComCode.WardId
'        Else
'           .ListPop GETSQLDEPT, lngTop, lngLeft ', objLisComCode.DeptCd
'        End If
        
'        txtCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
'        lblName.Caption = Trim(medGetP(.SelectedString, 2, ";"))
        If txtCd.Text <> "" Then Call GetTelephone
    End With
    Set objCodeList = Nothing
    
End Sub

Private Sub optDiv_Click(Index As Integer)
    Call FormClear
    If Index = 0 Then
        fraSave.Visible = True
        fraReview.Visible = False
        txtCd.SetFocus
    Else
        fraSave.Visible = False
        fraReview.Visible = True
        txtPtid.SetFocus
    End If
    
End Sub

Private Sub tabTelNo_Click()
    Dim RS          As Recordset
    Dim strWard     As String
    Dim strWardNm   As String
    Dim SSQL        As String
    
    
    On Error GoTo Errors
    Call medClearTable(tblward)
    strWard = tabTelNo.SelectedItem.Key
    strWardNm = tabTelNo.SelectedItem.Caption

    If strWard = "" Then Exit Sub
   
    SSQL = " SELECT distinct a." & F_PTID & " as ptid,c." & F_PTNM & " as ptnm," & _
           "         '' as tel,a." & F_PTWARDID & " as wardnm, a." & F_PTROOMID & " as roomid" & _
           " FROM " & T_HIS002 & " a," & T_LAB032 & " b," & T_HIS001 & " c" & _
           " WHERE " & _
                     DBW("a." & F_MAJDOCT & ">", "0") & _
           " AND " & DBW("a." & F_PTWARDID, strWard, 2) & _
           " AND " & F_BEDOUTDT2("a") & " is null " & _
           " AND a." & F_PTID & "=c." & F_PTID & _
           " AND " & DBJ("b.cdindex=* a." & F_PTWARDID) & _
           " AND not exists(SELECT * FROM " & T_LAB032 & " z " & _
           "                WHERE " & _
                            DBW("z.cdindex=", LC2_TelePhone) & _
           "        AND " & DBW("z.cdval1=", strWard) & _
           "        )"

     SSQL = SSQL & _
           " union all" & _
           " SELECT distinct a." & F_PTID & " as ptid,c." & F_PTNM & " as ptnm," & _
           "        b.field1 as tel,a." & F_PTWARDID & " as wardnm, a." & F_PTROOMID & " as roomid" & _
           " FROM " & T_HIS002 & " a," & T_LAB032 & " b," & T_HIS001 & " c" & _
           " WHERE " & _
                     DBW("a." & F_MAJDOCT & ">", "0") & _
           " AND " & DBW("a." & F_PTWARDID, strWard, 2) & _
           " AND " & F_BEDOUTDT2("a") & " is null " & _
           " AND a." & F_PTID & "=c." & F_PTID & _
           " AND " & DBW("b.cdindex=", LC2_TelePhone) & _
           " AND b.cdval1=a." & F_PTWARDID & _
           " ORDER BY roomid,ptid"

On Error GoTo Errors
    Set RS = New Recordset
    RS.Open SSQL, DBConn

    With tblward
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1: .RowHeight(.Row) = 13.63
                .Col = 1: .Value = strWardNm & "-" & RS.Fields("roomid").Value & ""
                .Col = 2: .Value = RS.Fields("ptid").Value & ""
                .Col = 3: .Value = RS.Fields("ptnm").Value & ""
                .Col = 4: .Value = RS.Fields("tel").Value & ""
                RS.MoveNext
            Loop
            .Row = 1: .Col = 1
            .Action = ActionActiveCell
        End If
    End With
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

Private Sub txtCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtCd_LostFocus()
'    Dim objData As clsBasisData
    Dim strData As String
    
    lblName.Caption = ""
    If txtCd.Text = "" Then Exit Sub
    
'    Set objData = New clsBasisData
    If optBussdiv(0).Value Then
        strData = GetWardNm(txtCd.Text)
        
        If strData = "" Then
            txtCd.Text = ""
            txtCd.SetFocus
        Else
            lblName.Caption = strData
        End If
        
'        If objLisComCode.WardId.Exists(Trim(txtCd.Text)) Then
'            objLisComCode.WardId.KeyChange Trim(txtCd.Text)
'            lblName.Caption = objLisComCode.WardId.Fields("wardnm")
'        Else
'            txtCd.Text = "": txtCd.SetFocus
'        End If
    Else
        strData = GetDeptNm(txtCd.Text)
        If strData = "" Then
            txtCd.Text = ""
            txtCd.SetFocus
        Else
            lblName.Caption = strData
        End If
        
'        If objLisComCode.DeptCd.Exists(Trim(txtCd.Text)) Then
'            objLisComCode.DeptCd.KeyChange Trim(txtCd.Text)
'            lblName.Caption = objLisComCode.DeptCd.Fields("deptnm")
'        Else
'            txtCd.Text = "": txtCd.SetFocus
'        End If
    End If
'    Set objData = Nothing
    Call GetTelephone
    
End Sub

Private Sub GetTelephone()
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", LC2_TelePhone) & _
           " AND " & DBW("cdval1=", Trim(txtCd.Text))
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        txtPhone.Text = RS.Fields("field1").Value & ""
    Else
        txtPhone.Text = ""
        txtPhone.SetFocus
    End If
    Set RS = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim SSQL    As String
    
    If txtCd.Text = "" Then
        MsgBox lblcode.Caption & " 정보가 누락되었습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    If txtPhone.Text = "" Then
        MsgBox "전화번호가 누락되었습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    If optBussdiv(0).Value Then
        SSQL = " delete " & T_LAB032 & _
               " WHERE " & _
                        DBW("cdindex=", LC2_TelePhone) & _
               " AND " & DBW("cdval1=", txtCd.Text)
        DBConn.Execute SSQL
        SSQL = "INSERT INTO " & T_LAB032 & "(cdindex,cdval1,field1,field2,field3) values(" & _
             DBV("cdindex", LC2_TelePhone, 1) & _
             DBV("cdval1", Trim(txtCd.Text), 1) & _
             DBV("field1", Trim(txtPhone.Text), 1) & _
             DBV("field2", lblName.Caption, 1) & _
             DBV("field3", "0") & ")"
    Else
        SSQL = " delete " & T_LAB032 & _
               " WHERE " & _
                        DBW("cdindex=", LC2_TelePhone) & _
               " AND " & DBW("cdval1=", txtCd.Text)
        DBConn.Execute SSQL
        SSQL = "INSERT INTO " & T_LAB032 & "(cdindex,cdval1,field1,field2,field3) values(" & _
                DBV("cdindex", LC2_TelePhone, 1) & _
                DBV("cdval1", Trim(txtCd.Text), 1) & _
                DBV("field1", Trim(txtPhone.Text), 1) & _
                DBV("field2", lblName.Caption, 1) & _
                DBV("field3", "1") & ")"
    End If
    DBConn.Execute SSQL
    Call TabIni
    DBConn.CommitTrans
    Exit Sub
SAVE_ERROR:
    MsgBox Err.Description
    DBConn.RollbackTrans
    
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtPtId_LostFocus()
    Dim objPat As clsPatient
    
    lblPtNm.Caption = "": lblsexage.Caption = "": lblbedindt.Caption = ""
    lbllocation.Caption = "": lblTestnm.Caption = "": lblphone.Caption = "": lblorddt.Caption = ""
    If txtPtid.Text = "" Then Exit Sub
    If IsNumeric(txtPtid.Text) Then txtPtid.Text = Format(txtPtid.Text, P_PatientIdFormat)
    Set objPat = New clsPatient
    With objPat
        If .GETPatient(txtPtid.Text) Then
            lblsexage.Caption = .SEXAGE
            lblbedindt.Caption = Format(.BedInDt, "####-##-##")
            lblPtNm.Caption = .ptnm
            lbldob.Caption = Format(.Dob, "####-##-##")
            If .INADMISSION Then
                optOption(1).Value = True
                
                lbllocation.Caption = .WardId & "-" & .ROOMID
            Else
                optOption(0).Value = True
            End If
            Call LastTestName
        End If
    End With
    Set objPat = Nothing
End Sub

Private Sub INPGetPhone(Optional ByVal DeptCd As String = "")
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim strWard     As String
    
    If DeptCd = "" Then
        strWard = medGetP(lbllocation.Caption, 1, "-")
        If strWard = "" Then Exit Sub
    Else
        strWard = DeptCd
    End If
    
    SSQL = " SELECT * FROM " & T_LAB032 & " WHERE " & _
                     DBW("cdindex=", LC2_TelePhone) & _
           " AND " & DBW("cdval1=", strWard)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        lblphone.Caption = RS.Fields("field1").Value & ""
    End If
    Set RS = Nothing
End Sub

Private Sub LastTestName()
    Dim SSQL As String
    Dim RS   As Recordset
'    Dim objData  As clsBasisData
    
'    Set objData = New clsBasisData
    
    SSQL = " SELECT i.ordcd,k.testnm,l.deptcd,l.orddt " & _
           " FROM " & T_LAB102 & " i," & T_LAB001 & " k, " & T_LAB101 & " l" & _
           " WHERE" & _
                    DBW("i.ptid=", txtPtid.Text) & _
           " AND i.orddt=(SELECT max(a.orddt) FROM " & T_LAB101 & " a " & _
                        " WHERE " & DBW("a.ptid=", txtPtid.Text) & " )" & _
           " AND i.ordno=(SELECT max(b.ordno) FROM " & T_LAB101 & " b " & _
           "              WHERE " & _
                                DBW("b.ptid=", txtPtid.Text) & _
           "               AND b.orddt=(SELECT max(orddt) " & _
                                      " FROM " & T_LAB101 & " c " & _
                                      " WHERE " & DBW("c.ptid=", txtPtid.Text) & "))" & _
           " AND i.ordcd=k.testcd" & _
           " AND i.ptid=l.ptid AND i.orddt=l.orddt AND i.ordno=l.ordno"

    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        lblorddt.Caption = Format(RS.Fields("orddt").Value & "", "####-##-##")
        lblTestnm.Caption = RS.Fields("testnm").Value & ""
        
        If optOption(0).Value Then
            Call INPGetPhone(RS.Fields("deptcd").Value & "")
'            objLisComCode.DeptCd.KeyChange RS.Fields("deptcd").Value & ""
            lbllocation.Caption = GetDeptNm(RS.Fields("deptcd").Value & "") 'objLisComCode.DeptCd.Fields("deptnm")
        Else
            Call INPGetPhone
        End If
    End If
    Set RS = Nothing
'    Set objData = Nothing
End Sub
