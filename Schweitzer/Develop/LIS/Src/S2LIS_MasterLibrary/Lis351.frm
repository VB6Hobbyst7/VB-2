VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm351ItemMaster 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사항목 마스터 등록"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis351.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   405
      Left            =   5595
      TabIndex        =   116
      Top             =   720
      Width           =   5475
      Begin MSComctlLib.TabStrip tabSpecimen 
         Height          =   300
         Left            =   60
         TabIndex        =   117
         Top             =   75
         Width           =   4950
         _ExtentX        =   8731
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H0085A3A3&
         BorderWidth     =   2
         Height          =   345
         Left            =   45
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   4995
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   405
      Left            =   105
      TabIndex        =   114
      Top             =   720
      Width           =   5475
      Begin MSComctlLib.TabStrip tabItem 
         Height          =   300
         Left            =   60
         TabIndex        =   115
         Top             =   75
         Width           =   5340
         _ExtentX        =   9419
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
         BorderColor     =   &H0086A0A2&
         BorderWidth     =   2
         Height          =   345
         Left            =   45
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   5385
      End
   End
   Begin VB.TextBox txtWorkUnit 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00F1F5F4&
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
      Left            =   1485
      TabIndex        =   111
      Top             =   8460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00DBE6E6&
      Height          =   750
      Left            =   165
      TabIndex        =   47
      Top             =   -30
      Width           =   10470
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F4F0F2&
         Caption         =   "<< (&P)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   4785
         Style           =   1  '그래픽
         TabIndex        =   64
         Tag             =   "124"
         Top             =   165
         Width           =   1320
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F4F0F2&
         Caption         =   "(&N) >>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   6105
         Style           =   1  '그래픽
         TabIndex        =   63
         Tag             =   "124"
         Top             =   165
         Width           =   1320
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
         Height          =   315
         Left            =   2595
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis351.frx":038A
         Style           =   1  '그래픽
         TabIndex        =   62
         Top             =   195
         Width           =   300
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
         Left            =   7755
         Style           =   1  '그래픽
         TabIndex        =   50
         Tag             =   "124"
         Top             =   165
         Visible         =   0   'False
         Width           =   1320
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
         Height          =   510
         Left            =   9075
         Style           =   1  '그래픽
         TabIndex        =   49
         Tag             =   "128"
         Top             =   165
         Width           =   1320
      End
      Begin VB.TextBox txtTestCd 
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
         Height          =   330
         Left            =   1155
         TabIndex        =   0
         Top             =   195
         Width           =   1425
      End
      Begin VB.Label lblItemCd 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   48
         Tag             =   "35121"
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7665
      Left            =   150
      TabIndex        =   70
      Top             =   1005
      Width           =   5385
      Begin VB.CheckBox chkNot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방불가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3510
         TabIndex        =   119
         Tag             =   "35102"
         Top             =   1635
         Width           =   1530
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0CFC2&
         Caption         =   "삭제(&D)"
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
         Left            =   4155
         Style           =   1  '그래픽
         TabIndex        =   33
         Tag             =   "35301"
         Top             =   675
         Width           =   1155
      End
      Begin VB.TextBox txtRptSeq 
         Alignment       =   2  '가운데 맞춤
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
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   21
         Top             =   5865
         Width           =   675
      End
      Begin VB.TextBox txtItemSeq 
         Alignment       =   2  '가운데 맞춤
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
         Height          =   315
         Left            =   4260
         MaxLength       =   4
         TabIndex        =   22
         Top             =   5880
         Width           =   645
      End
      Begin VB.PictureBox picRstDiv 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1665
         ScaleHeight     =   300
         ScaleWidth      =   3390
         TabIndex        =   75
         Top             =   4605
         Width           =   3450
         Begin VB.OptionButton optRstDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "Alternative"
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
            Index           =   1
            Left            =   1740
            TabIndex        =   17
            Tag             =   "35135"
            Top             =   15
            Width           =   1350
         End
         Begin VB.OptionButton optRstDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "Required"
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
            Index           =   0
            Left            =   75
            TabIndex        =   16
            Tag             =   "35136"
            Top             =   15
            Width           =   1560
         End
      End
      Begin VB.ComboBox cboRefLab 
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
         Height          =   300
         ItemData        =   "Lis351.frx":0914
         Left            =   1665
         List            =   "Lis351.frx":0921
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   5370
         Width           =   3495
      End
      Begin VB.ComboBox cboRstType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Lis351.frx":093E
         Left            =   1665
         List            =   "Lis351.frx":094B
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   5010
         Width           =   2235
      End
      Begin VB.PictureBox picTestDiv 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1665
         ScaleHeight     =   300
         ScaleWidth      =   3390
         TabIndex        =   73
         Top             =   3420
         Width           =   3450
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "혈액형"
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
            Index           =   3
            Left            =   2415
            TabIndex        =   113
            Tag             =   "35136"
            Top             =   15
            Width           =   840
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "미생물"
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
            Index           =   2
            Left            =   1515
            TabIndex        =   9
            Tag             =   "35136"
            Top             =   15
            Width           =   840
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "특수"
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
            Index           =   1
            Left            =   795
            TabIndex        =   8
            Tag             =   "35135"
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00F1F5F4&
            Caption         =   "일반"
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
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Tag             =   "35136"
            Top             =   15
            Width           =   660
         End
      End
      Begin VB.ComboBox cboWorkArea 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Lis351.frx":0962
         Left            =   1665
         List            =   "Lis351.frx":0964
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   2925
         Width           =   3405
      End
      Begin VB.TextBox txtFullNm 
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
         Height          =   315
         Left            =   1650
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2340
         Width           =   3405
      End
      Begin VB.TextBox txtAbbrNm10 
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
         Height          =   330
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1980
         Width           =   3405
      End
      Begin VB.TextBox txtAbbrNm5 
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
         Height          =   330
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1620
         Width           =   1755
      End
      Begin VB.CheckBox chkGrpFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Graphic결과 여부"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   285
         TabIndex        =   26
         Tag             =   "35102"
         Top             =   7035
         Width           =   1995
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "수정(&E)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2985
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "35106"
         Top             =   255
         Width           =   1155
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1815
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   29
         Tag             =   "35106"
         Top             =   255
         Width           =   1155
      End
      Begin VB.CheckBox chkDetailFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "상세항목"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2385
         TabIndex        =   27
         Tag             =   "35102"
         Top             =   7065
         Width           =   1110
      End
      Begin VB.ComboBox cboAttrCd 
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
         Height          =   300
         ItemData        =   "Lis351.frx":0966
         Left            =   1680
         List            =   "Lis351.frx":0968
         Style           =   2  '드롭다운 목록
         TabIndex        =   24
         Top             =   6615
         Width           =   1695
      End
      Begin VB.PictureBox picPanelFg 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1665
         ScaleHeight     =   300
         ScaleWidth      =   3390
         TabIndex        =   72
         Top             =   3810
         Width           =   3450
         Begin VB.OptionButton optPanFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "상세검사"
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
            Index           =   2
            Left            =   2175
            TabIndex        =   12
            Tag             =   "35136"
            Top             =   15
            Width           =   1095
         End
         Begin VB.OptionButton optPanFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "그룹처방"
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
            Index           =   1
            Left            =   990
            TabIndex        =   11
            Tag             =   "35135"
            Top             =   15
            Width           =   1095
         End
         Begin VB.OptionButton optPanFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "없음"
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
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Tag             =   "35135"
            Top             =   15
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdClinicalNote 
         BackColor       =   &H00F4F0F2&
         Caption         =   "C&linical Notice"
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
         Left            =   3585
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   25
         Tag             =   "35106"
         Top             =   6390
         Width           =   1515
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소(&U)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4155
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   31
         Tag             =   "35106"
         Top             =   255
         Width           =   1155
      End
      Begin VB.PictureBox picTextFg 
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   1665
         ScaleHeight     =   300
         ScaleWidth      =   3390
         TabIndex        =   71
         Top             =   4200
         Width           =   3450
         Begin VB.OptionButton optTextFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "없음"
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
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Tag             =   "35135"
            Top             =   30
            Width           =   780
         End
         Begin VB.OptionButton optTextFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "TextOnly"
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
            Index           =   1
            Left            =   960
            TabIndex        =   14
            Tag             =   "35136"
            Top             =   15
            Width           =   1095
         End
         Begin VB.OptionButton optTextFg 
            BackColor       =   &H00F1F5F4&
            Caption         =   "일반&&Text"
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
            Index           =   2
            Left            =   2175
            TabIndex        =   15
            Tag             =   "35136"
            Top             =   15
            Width           =   1290
         End
      End
      Begin VB.CommandButton cmd_S_RstType 
         BackColor       =   &H00F4F0F2&
         Caption         =   "특수검사유형"
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
         Left            =   3915
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   19
         Tag             =   "35106"
         Top             =   5010
         Width           =   1215
      End
      Begin VB.CommandButton cmdDetailItem 
         BackColor       =   &H00F4F0F2&
         Caption         =   "상세항목설정(&D)"
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
         Left            =   3570
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "35106"
         Top             =   6915
         Width           =   1515
      End
      Begin VB.ComboBox cboStGroup 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Lis351.frx":096A
         Left            =   1680
         List            =   "Lis351.frx":0977
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   6270
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpApplyDate 
         Height          =   330
         Left            =   1650
         TabIndex        =   1
         Top             =   825
         Width           =   1785
         _ExtentX        =   3149
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   16711683
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpExpireDate 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
         Top             =   1215
         Width           =   1785
         _ExtentX        =   3149
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd"
         DateIsNull      =   -1  'True
         Format          =   16711683
         CurrentDate     =   36328
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2355
         TabIndex        =   74
         Top             =   5850
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "picRstDiv"
         BuddyDispid     =   196626
         OrigLeft        =   1830
         OrigTop         =   3765
         OrigRight       =   2070
         OrigBottom      =   4110
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   4906
         TabIndex        =   76
         Top             =   5880
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "UpDown1"
         BuddyDispid     =   196705
         OrigLeft        =   1830
         OrigTop         =   3765
         OrigRight       =   2070
         OrigBottom      =   4110
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과구분 : "
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
         TabIndex        =   92
         Tag             =   "35123"
         Top             =   4695
         Width           =   900
      End
      Begin VB.Label lblRptSeq 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고서순서 : "
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
         TabIndex        =   91
         Tag             =   "35125"
         Top             =   5925
         Width           =   1080
      End
      Begin VB.Label lblRstTp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과유형 : "
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
         TabIndex        =   90
         Tag             =   "35126"
         Top             =   5085
         Width           =   900
      End
      Begin VB.Label lblRefLab 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "외부검사기관 : "
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
         Left            =   240
         TabIndex        =   89
         Tag             =   "35124"
         Top             =   5445
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사구분 : "
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
         TabIndex        =   88
         Tag             =   "35123"
         Top             =   3525
         Width           =   900
      End
      Begin VB.Label lblWrkArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Work Area : "
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
         TabIndex        =   87
         Tag             =   "35133"
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label lblFullNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사명 : "
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
         TabIndex        =   86
         Tag             =   "35120"
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label lblAbbNm2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "약어2 (10문자) : "
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
         Left            =   240
         TabIndex        =   85
         Tag             =   "35113"
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label lblAbbNm1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "약어1 (5문자) : "
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
         TabIndex        =   84
         Tag             =   "35112"
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblExpDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폐기일 : "
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
         TabIndex        =   83
         Tag             =   "35118"
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lblAppDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적용일 : "
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
         TabIndex        =   82
         Tag             =   "35114"
         Top             =   915
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "속성코드 : "
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
         Left            =   240
         TabIndex        =   81
         Tag             =   "35126"
         Top             =   6690
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Panel 처방 : "
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
         TabIndex        =   80
         Tag             =   "35124"
         Top             =   3900
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "텍스트결과 : "
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
         TabIndex        =   79
         Tag             =   "35126"
         Top             =   4305
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "통계출력순서 : "
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
         Left            =   2985
         TabIndex        =   78
         Tag             =   "35125"
         Top             =   5940
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "통계 Group : "
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
         TabIndex        =   77
         Tag             =   "35133"
         Top             =   6345
         Width           =   1110
      End
      Begin VB.Label Label22 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사항목정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   195
         Left            =   270
         TabIndex        =   93
         Top             =   360
         Width           =   1290
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00DDF0F5&
         FillStyle       =   0  '단색
         Height          =   390
         Index           =   0
         Left            =   105
         Shape           =   4  '둥근 사각형
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7665
      Left            =   5640
      TabIndex        =   32
      Top             =   1005
      Width           =   5025
      Begin VB.CommandButton cmdRefer 
         BackColor       =   &H00F4F0F2&
         Caption         =   "참고치 등록(&R)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3435
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   57
         Tag             =   "35106"
         Top             =   6795
         Width           =   1455
      End
      Begin VB.TextBox txtStoreCd 
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3735
         Width           =   945
      End
      Begin VB.CommandButton cmdSpecimen 
         BackColor       =   &H00F4F0F2&
         Caption         =   "검체 등록(&S)"
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
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   51
         Tag             =   "35109"
         Top             =   225
         Width           =   1320
      End
      Begin VB.CheckBox chkRndFg 
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         Height          =   390
         Left            =   345
         TabIndex        =   37
         Tag             =   "35104"
         Top             =   5370
         Width           =   210
      End
      Begin VB.CheckBox chkStatFg 
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         Height          =   420
         Left            =   2055
         TabIndex        =   36
         Tag             =   "35105"
         Top             =   5355
         Width           =   195
      End
      Begin VB.CheckBox chkPanicFg 
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   35
         Tag             =   "35103"
         Top             =   5880
         Width           =   225
      End
      Begin VB.CheckBox chkDeltaFg 
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   34
         Tag             =   "35101"
         Top             =   6255
         Width           =   195
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   300
         Left            =   2025
         TabIndex        =   95
         Top             =   750
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblSpcName 
         Height          =   285
         Left            =   2025
         TabIndex        =   96
         Top             =   1125
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
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
      Begin MedControls1.LisLabel lblSpcQty 
         Height          =   300
         Left            =   2040
         TabIndex        =   97
         Top             =   4080
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblSpcAppDt 
         Height          =   300
         Left            =   2040
         TabIndex        =   98
         Top             =   2295
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblSpcExpDt 
         Height          =   300
         Left            =   2040
         TabIndex        =   99
         Top             =   2655
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblRstUnit 
         Height          =   300
         Left            =   2040
         TabIndex        =   100
         Top             =   3015
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblAvalVal 
         Height          =   300
         Left            =   2040
         TabIndex        =   101
         Top             =   3375
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblTestCost 
         Height          =   300
         Left            =   2040
         TabIndex        =   102
         Top             =   4455
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblTatAvg 
         Height          =   300
         Left            =   2040
         TabIndex        =   103
         Top             =   4815
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblSpcUnit 
         Height          =   300
         Left            =   3885
         TabIndex        =   104
         Top             =   4110
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblPanicFrVal 
         Height          =   300
         Left            =   1905
         TabIndex        =   105
         Top             =   5850
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblPanicToVal 
         Height          =   300
         Left            =   3525
         TabIndex        =   106
         Top             =   5850
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblDeltaVal1 
         Height          =   300
         Left            =   2115
         TabIndex        =   107
         Top             =   6225
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblDeltaVal2 
         Height          =   300
         Left            =   3750
         TabIndex        =   108
         Top             =   6240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblLabelCnt 
         Height          =   300
         Left            =   2430
         TabIndex        =   109
         Top             =   6720
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
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
      Begin MedControls1.LisLabel lblSpcGrpCd 
         Height          =   300
         Left            =   2430
         TabIndex        =   110
         Top             =   7095
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
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
      Begin MSComctlLib.TabStrip tabAppDt 
         Height          =   300
         Left            =   555
         TabIndex        =   118
         Top             =   1830
         Width           =   4170
         _ExtentX        =   7355
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
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   360
         Left            =   540
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label23 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   195
         Left            =   480
         TabIndex        =   94
         Top             =   330
         Width           =   870
      End
      Begin VB.Label Label21 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "(-)"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1815
         TabIndex        =   69
         Top             =   6300
         Width           =   285
      End
      Begin VB.Label Label20 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3450
         TabIndex        =   68
         Top             =   6315
         Width           =   285
      End
      Begin VB.Label Label19 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "%"
         Height          =   270
         Left            =   3165
         TabIndex        =   67
         Top             =   6270
         Width           =   285
      End
      Begin VB.Label lblCnt2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Barcode 출력장수 : "
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
         Left            =   630
         TabIndex        =   66
         Tag             =   "35211"
         Top             =   6825
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체군 : "
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
         Left            =   1575
         TabIndex        =   65
         Tag             =   "35126"
         Top             =   7170
         Width           =   720
      End
      Begin VB.Label Label13 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Delta Check"
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
         Left            =   630
         TabIndex        =   61
         Top             =   6285
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Panic Check"
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
         Left            =   630
         TabIndex        =   60
         Top             =   5910
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급 검사 여부"
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
         Left            =   2325
         TabIndex        =   59
         Top             =   5460
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "일괄 채혈 여부"
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
         Left            =   600
         TabIndex        =   58
         Top             =   5475
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사소요시간 : "
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
         Left            =   615
         TabIndex        =   56
         Tag             =   "35116"
         Top             =   4875
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수가 : "
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
         Left            =   615
         TabIndex        =   55
         Tag             =   "35116"
         Top             =   4515
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "단위 : "
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
         Left            =   3360
         TabIndex        =   54
         Tag             =   "35116"
         Top             =   4170
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체량 : "
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
         Left            =   615
         TabIndex        =   52
         Tag             =   "35116"
         Top             =   4155
         Width           =   720
      End
      Begin VB.Label lblSpcCdL 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체코드 : "
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
         Left            =   600
         TabIndex        =   46
         Tag             =   "35127"
         Top             =   810
         Width           =   900
      End
      Begin VB.Label lblSpcNmL 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 체  명 : "
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
         Left            =   600
         TabIndex        =   45
         Tag             =   "35128"
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과단위 : "
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
         Left            =   615
         TabIndex        =   44
         Tag             =   "35132"
         Top             =   3075
         Width           =   900
      End
      Begin VB.Label lblDecPnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "유효숫자 : "
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
         Left            =   615
         TabIndex        =   43
         Tag             =   "35116"
         Top             =   3420
         Width           =   900
      End
      Begin VB.Label Label15 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "%"
         Height          =   270
         Left            =   4755
         TabIndex        =   42
         Top             =   6270
         Width           =   285
      End
      Begin VB.Label Label16 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   270
         Left            =   3225
         TabIndex        =   41
         Top             =   5910
         Width           =   285
      End
      Begin VB.Label lblAppDt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적용일 : "
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
         Left            =   615
         TabIndex        =   40
         Tag             =   "35115"
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label lblExpDt1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폐기일 : "
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
         Left            =   615
         TabIndex        =   39
         Tag             =   "35119"
         Top             =   2715
         Width           =   720
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보관방법 : "
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
         Left            =   615
         TabIndex        =   38
         Tag             =   "35122"
         Top             =   3780
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00DDF0F5&
         FillStyle       =   0  '단색
         Height          =   390
         Index           =   1
         Left            =   225
         Shape           =   4  '둥근 사각형
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.Label lblWrkUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Workload Unit"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   112
      Tag             =   "35134"
      Top             =   8550
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frm351ItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private MySqlStmt As New clsLISSqlStatement ' SQL 클래스
Private MyItems As New clsItems             ' 검사항목 클래스
Private MyItem As New clsItem               ' 검사항목 클래스
Private MySpecimens As New clsSpecimens     ' 검체 클래스

Private InsertFlag As Integer
Private UpdateFlag As Integer

Private SvApplyDt As String

'Private Sub objCodeList_LostFocus()
'    txtTestCd.SetFocus
'End Sub
'
Private Sub cmd_S_RstType_Click()
    With frm366EDefine
        Call SetParent(.hWnd, gParentWhnd)
        .Show
        '.txtTestCd.Text = txtTestCd.Text
        'Call .Raise_TestCd_Keypress
        .ZOrder 0
    End With
End Sub

Private Sub cmdCancel_Click()

    Call CancelRoutine
    'If tabItem.Pages.Count > 0 Then tabItem.Value = 0: Call tabItem_Click(0) 'tabItem.Pages(1).Selected = True
   If TabItem.Tabs.Count > 0 Then TabItem.Tabs(1).Selected = True
    
End Sub

Private Sub CancelRoutine()
    
    If Not ConfirmExit Then Exit Sub
    
    InsertFlag = 0
    UpdateFlag = 0
    
    Call LockRtn(1, True)
    
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdNew.Caption = "추가"
    cmdEdit.Caption = "수정"
    cmdCancel.Enabled = False

End Sub
Private Sub cmdClear_Click()

    If Not ConfirmExit Then Exit Sub
    Call ClearRtn(5)
    txtTestCd.Text = ""
    txtTestCd.SetFocus

End Sub

Private Sub cmdClinicalNote_Click()
   
    With frm343Template
        Call SetParent(.hWnd, gParentWhnd)
        .Rkey = LC4_ClinicalNotice
        .RName = "Clinical Notice"
        .Show
        .ZOrder
    End With
    
End Sub

Private Sub cmdDelete_Click()
    
    Dim Resp As VbMsgBoxResult
    
    'If tabItem.Pages.Count <= 0 Then Exit Sub
    If TabItem.Tabs.Count <= 0 Then Exit Sub
    
    Resp = MsgBox("해당 적용일의 데이타를 모두 삭제하시겠습니까?", vbQuestion, "검사항목 등록")
    If Resp = vbNo Then Exit Sub
    
    With MyItem
       Call Lab001Move(MyItem)
       .ItemDelete
       MyItems.Remove Format(dtpApplyDate, CS_DateDbFormat)
    End With
    If lstItemList.Exists(txtTestCd.Text) Then
        lstItemList.KeyChange txtTestCd.Text
        lstItemList.Delete
    End If

    txtTestCd.Text = ""
    Call cmdClear_Click
    
End Sub

Private Sub cmdDetailItem_Click()
    
    With frm341Common1
         Call SetParent(.hWnd, gParentWhnd)
        .Rkey = LC2_Detail
        .RName = "상세항목 관리"
        .txtFIndex.Text = txtTestCd.Text
        Call .Raise_lstSKey1_MouseUp
        .Show
        .ZOrder 0
    End With

End Sub

'% Edit Button 클릭 : Data Update

Private Sub cmdEdit_Click()

    If UpdateFlag = 1 Then  ' Update
        cmdEdit.Caption = "수정"
        With MyItem
           Call Lab001Move(MyItem)
           .ItemUpdate
           MyItems.Update Format(dtpApplyDate, CS_DateDbFormat), MyItem
        End With
        UpdateFlag = 0
        Call LockRtn(1, True)
        cmdNew.Enabled = True
        cmdCancel.Enabled = False
       
    Else    ' Edit
        dtpApplyDate.Enabled = False
        cmdEdit.Caption = "저장"
        UpdateFlag = 1
        Call LockRtn(1, False)
        cmdNew.Enabled = False
        cmdCancel.Enabled = True
    End If

End Sub

'% 종료
Private Sub cmdExit_Click()
    If Not ConfirmExit Then Exit Sub
    Unload Me
End Sub


Private Sub cmdFind_Click(Index As Integer)

    Dim i As Integer
    
    If txtTestCd.Text = "" Then Exit Sub
    If Not ConfirmExit Then Exit Sub

'    I = medListFind(lstItemList, txtTestCd.Text)
    If Not lstItemList.Exists(txtTestCd.Text) Then Exit Sub
    Call lstItemList.KeyChange(txtTestCd.Text)

'    If I < 0 Then Exit Sub
    Select Case Index
        Case 0:   'Previous
            'If I <= 0 Then Exit Sub
            'txtTestCd.Text = lstItemList.List(I - 1)
            lstItemList.MovePrevious
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
        Case 1:   'Next
'            If I >= lstItemList.ListCount - 1 Then Exit Sub
'            txtTestCd.Text = lstItemList.List(I + 1)
            lstItemList.MoveNext
            If lstItemList.EOF Or lstItemList.Key = "" Then Exit Sub
            txtTestCd.Text = lstItemList.Key
    End Select
    Call txtTestCd_KeyPress(vbKeyReturn)

End Sub

'% New Button 클릭 : Data Insert

Private Sub cmdNew_Click()

    If InsertFlag = 1 Then ' Insert
        If SvApplyDt <> "" And SvApplyDt >= Format(dtpApplyDate.Value, CS_DateDbFormat) Then
            MsgBox "적용일을 수정하세요.."
            dtpApplyDate.SetFocus
            Exit Sub
        End If
        If Trim(txtFullNm.Text) = "" Or Trim(txtAbbrNm5.Text) = "" Or Trim(txtAbbrNm10.Text) = "" Then
            MsgBox "검사명(약어명)을 모두 입력하세요.", vbInformation, "메세지"
            Exit Sub
        End If
        If cboWorkArea.ListIndex < 0 Then
            MsgBox "Work Area를 입력하세요", vbInformation, "메세지"
            cboWorkArea.SetFocus
            Exit Sub
        End If
        cmdNew.Caption = "추가"
        With MyItem
            Call Lab001Move(MyItem)
            .ItemInsert
            
            '성바오로병원 OCS처방 전달용임
            
            MyItems.Add Format(dtpApplyDate, CS_DateDbFormat), MyItem
        End With
      
        'New Item이면 List에 추가...
'        I = medListFind(lstItemList, txtTestCd.Text)
'        If txtTestCd.Text <> lstItemList.List(I) Then lstItemList.AddItem txtTestCd.Text
        If Not lstItemList.Exists(txtTestCd.Text) Then
            lstItemList.Sort = False
            lstItemList.AddNew txtTestCd.Text, "New Item"
            lstItemList.Sort = True
        End If
        
        InsertFlag = 0
        txtTestCd_KeyPress (vbKeyReturn)
        Call LockRtn(1, True)
        

        
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False

    Else    ' New
        cmdNew.Caption = "저장"
        InsertFlag = 1
        Call ClearRtn(1)
        Call LockRtn(1, False)
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        'If tabItem.Pages.Count > 0 Then
        If TabItem.Tabs.Count > 0 Then
            SvApplyDt = Format(dtpApplyDate.Value, CS_DateDbFormat)
        Else
            SvApplyDt = ""
        End If
        dtpApplyDate.Value = Format(Now, CS_DateLongFormat)
        dtpApplyDate.SetFocus
    End If

End Sub


'% List 버튼을 클릭한 경우 코드리스트를 팝업한다.
Private Sub cmdPopupList_Click()

    Dim tmpSql As String
    Dim lngTop As Long, lngLeft As Long

    If Not ConfirmExit Then Exit Sub

    Set objCodeList = New clsPopUpList
    With objCodeList
        .AutoGap = True
        .FormWidth = 6500
        lngTop = txtTestCd.Top + 2350
        lngLeft = Me.Left + Frame1.Left + txtTestCd.Left + 50
        .Connection = DBConn
        .Tag = "TestCd"
        .Delimiter = ";"
        .FormCaption = "검사항목 리스트"
        .ColumnHeaderText = "검사코드;검사명;약어5;약어10"
        tmpSql = MySqlStmt.SqlLAB001CodeList
        '.ListPop tmpSql, lngTop, lngLeft
        .LoadPopUp tmpSql ' , lngTop, lngLeft,  lstItemList
        txtTestCd.Text = Trim(medShift(.SelectedString, ";"))
        Call txtTestCd_KeyPress(vbKeyReturn)
    End With

End Sub


' Reference(기준치) 등록 창 Popup
Private Sub cmdRefer_Click()
    With frm353Reference
        .Show
        Call SetParent(.hWnd, gParentWhnd)
        .txtTestCd.Text = txtTestCd.Text
        Call .Raise_TestCd_Keypress
        .cboSpcCd.ListIndex = medComboFind(.cboSpcCd, lblSpcCd.Caption)
        DoEvents
        .ZOrder 0
    End With
End Sub

' Specimen(검체) 등록 창 Popup
Private Sub cmdSpecimen_Click()
    With frm352Specimen
        Call SetParent(.hWnd, gParentWhnd)
        .Show
        .txtTestCd.Text = txtTestCd.Text
        Call .Raise_TestCd_Keypress
        .ZOrder 0
    End With
End Sub

Private Sub Form_Activate()
    If Me.Visible Then txtTestCd.SetFocus
End Sub

Private Sub Form_Deactivate()
    Set objCodeList = Nothing
End Sub

Private Sub Form_Load()

'    tabItem.Pages.Clear
'    tabSpecimen.Pages.Clear

    TabItem.Tabs.Clear
    tabSpecimen.Tabs.Clear

'    Me.HelpContextID = HLP_ItemMaster

    TabItem.ZOrder 0
    tabSpecimen.ZOrder 0
    dtpApplyDate.Value = Format(Now, "YYYY-MM-DD")
    dtpExpireDate.Value = ""

    InsertFlag = 0
    UpdateFlag = 0
    
  
    chkNot.Visible = False
    
    Call MyItem.GetWorkArea(cboWorkArea): DoEvents
    Call MyItem.GetGroupCd(cboStGroup): DoEvents
    Call MyItem.GetOutLabList(cboRefLab): DoEvents
    Call MyItem.GetItemList(lstItemList): DoEvents

    cmdDelete.Enabled = ObjMyUser.isdeveloper ' ObjSysInfo.IsDeveloper 'gIsDeveloper

End Sub



Private Sub Form_Unload(Cancel As Integer)

    Set objCodeList = Nothing
    Set MySqlStmt = Nothing
    Set MyItems = Nothing
    Set MyItem = Nothing
    Set MySpecimens = Nothing

End Sub

Private Sub optPanFg_Click(Index As Integer)
    Select Case Index
        Case 0:   '일반
            chkDetailFg.Enabled = True
            chkGrpFg.Enabled = True
            cmdDetailItem.Visible = False
            picTextFg.Enabled = True
            picRstDiv.Enabled = True
            cboRstType.Enabled = True
            cboRefLab.Enabled = True
        Case 1:   '그룹
            chkDetailFg.Value = 0
            chkDetailFg.Enabled = False
            chkGrpFg.Value = 0
            chkGrpFg.Enabled = False
            cmdDetailItem.Visible = False
            optTextFg(0).Value = True
            picTextFg.Enabled = False
            optRstDiv(0).Value = True
            picRstDiv.Enabled = False
            If cboRstType.ListCount > 0 Then
                cboRstType.ListIndex = 0
            Else
                cboRstType.ListIndex = -1
            End If
            cboRstType.Enabled = False
            cboRefLab.Enabled = False
        Case 2:    '상세
            chkDetailFg.Value = 0
            chkDetailFg.Enabled = False
            chkGrpFg.Value = 0
            chkGrpFg.Enabled = True
            cmdDetailItem.Visible = True
            picTextFg.Enabled = True
            picRstDiv.Enabled = True
            cboRstType.Enabled = True
            cboRefLab.Enabled = True
     End Select
End Sub

Private Sub optTestDiv_Click(Index As Integer)
    Select Case Index
    Case 0, 3:
        Call MyItem.GetRstType(cboRstType, "0")
        cboRstType.Enabled = True
        picPanelFg.Enabled = True
        cmd_S_RstType.Visible = False
        chkDetailFg.Enabled = True
    Case 1:
        Call MyItem.GetRstType(cboRstType, "1")
        cboRstType.Enabled = False
        optPanFg(0).Value = True
        picPanelFg.Enabled = False
        cmd_S_RstType.Visible = True
        chkDetailFg.Enabled = False
    Case 2:
        Call MyItem.GetRstType(cboRstType, "2")
        cboRstType.Enabled = True
        optPanFg(0).Value = True
        picPanelFg.Enabled = True
        cmd_S_RstType.Visible = False
        chkDetailFg.Enabled = True
    End Select
    '특수검사를 제외한 검사유형에서는 검사결과가 Text only 일수가 없다.
    Select Case Index
        Case 0, 2, 3
            optTextFg(1).Enabled = False
        Case Else
            optTextFg(1).Enabled = True
    End Select
    
End Sub


'% 검체의 적용일자(tabAppDt)를 클릭하면 검체의 상세정보 Display
Private Sub tabAppDt_Click()

    Dim tmpStr As String

    'If tabAppDt.Pages.Count <= 0 Then Exit Sub
    If tabAppDt.Tabs.Count <= 0 Then Exit Sub
    
    'tmpStr = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
    tmpStr = Format(tabAppDt.SelectedItem.Caption, CS_DateDbFormat)
    With MySpecimens.Specimen(tmpStr)
        txtTestCd.Text = .TestCd
        lblSpcCd.Caption = .SpcCd
        lblSpcAppDt.Caption = Format(.ApplyDt, CS_DateMask)
        lblSpcGrpCd.Caption = .SpcGrpCd
        lblLabelCnt.Caption = .LabelCnt
        lblRstUnit.Caption = .RstUnit
        chkRndFg.Value = Val(.RndFg)
        chkStatFg.Value = Val(.StatFg)
        lblAvalVal.Caption = .AvalVal
        chkPanicFg.Value = Val(.PanicFg)
        lblPanicFrVal.Caption = .PanicFrVal
        lblPanicToVal.Caption = .PanicToVal
        chkDeltaFg.Value = Val(.DeltaFg)
        lblDeltaVal1.Caption = .DeltaVal1
        lblDeltaVal2.Caption = .DeltaVal2
        lblTestCost.Caption = .TestCost
        txtStoreCd.Text = .StoreCd
        lblTatAvg.Caption = .TatAvg
        lblSpcQty.Caption = .SpcQty
        lblSpcUnit.Caption = .SpcUnit
        lblSpcExpDt.Caption = Format(.ExpDt, CS_DateMask)
    End With

End Sub


Private Sub tabItem_Click()


    Dim tmpStr As String

    Call CancelRoutine
    
    'tmpStr = Format(tabItem.SelectedItem.Caption, CS_DateDbFormat)
    tmpStr = Format(TabItem.SelectedItem.Caption, CS_DateDbFormat)
    With MyItems.Item(tmpStr)
        '적용일
        dtpApplyDate.Value = Format(.ApplyDt, CS_DateMask)
        '폐기일
        If Trim(.ExpDt) <> "" Then
            dtpExpireDate.Value = Format(.ExpDt, CS_DateMask)
        Else
            dtpExpireDate.Value = ""
        End If
        '검사명/약어(5)/약어(10)
        txtAbbrNm5.Text = .AbbrNm5
        txtAbbrNm10.Text = .AbbrNm10
        txtFullNm.Text = .TestNm
        '검사구분 : 0-일반검사, 1-기타검사, 2-미생물검사
        optTestDiv(Val(.TestDiv)).Value = True
        'WorkArea
        cboWorkArea.ListIndex = medComboFind(cboWorkArea, .WorkArea)
        If medGetP(cboWorkArea.List(cboWorkArea.ListIndex), 1, " ") <> .WorkArea Then cboWorkArea.ListIndex = -1
        '결과유형 : Null-일반, R-Ratio, F- Free
        cboRstType.ListIndex = medComboFind(cboRstType, .RstType)
        If .RstType = "" And cboRstType.ListIndex > 0 Then cboRstType.ListIndex = 0
        '레포트Seq / Workload Unit
        txtRptSeq.Text = .RptSeq
        txtWorkUnit.Text = .WorkUnit
        '외부의뢰기관
        cboRefLab.ListIndex = medComboFind(cboRefLab, .OutLabCd)
        If Trim(medGetP(cboRefLab.List(cboRefLab.ListIndex), 1, " ")) <> .OutLabCd Then cboRefLab.ListIndex = -1
        '속성코드
        cboAttrCd.ListIndex = medComboFind(cboAttrCd, .AttrCd)
        If medGetP(cboAttrCd.List(cboAttrCd.ListIndex), 1, " ") <> .AttrCd Then cboAttrCd.ListIndex = -1
        'Panel Fg : Null-일반, G-그룹처방, D-상세처방
        If .PanelFg = "" Then
            optPanFg(0).Value = True
        ElseIf .PanelFg = "G" Then
            optPanFg(1).Value = True
        ElseIf .PanelFg = "D" Then
            optPanFg(2).Value = True
        End If
        '결과구분 : R-Required, A-Alternative
        If .RstDiv = "R" Then
            optRstDiv(0).Value = True
        ElseIf .RstDiv = "A" Then
            optRstDiv(1).Value = True
        End If
        'Text결과유형 : 0-없음, 1-Text Only, 2- 일반&Text
        optTextFg(Val(.TxtType)).Value = True
        chkGrpFg.Value = Val(.GrpFg)
        'Detail Fg : Null-일반항목, '*'-상세항목
        If .DetailFg = "*" Then
            chkDetailFg.Value = 1
        Else
            chkDetailFg.Value = 0
        End If
        '통계Seq
        txtItemSeq.Text = .ItemSeq
        '통계GroupCd
        cboStGroup.ListIndex = medComboFind(cboStGroup, .GroupCd)
        If medGetP(cboStGroup.List(cboStGroup.ListIndex), 1, " ") <> .GroupCd Then cboStGroup.ListIndex = -1
        
    End With
    Call LockRtn(1, True)

    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False

End Sub


Private Sub tabSpecimen_Click()

    Dim tmpStr As String
    Dim tmpSql As String
    Dim tmpSpc As String, tmpSeq As String

'    tmpSpc = Mid(tabSpecimen.SelectedItem.Name, 2)
'    tmpSeq = CStr(tabSpecimen.SelectedItem.Index + 1)
'    tmpSql = MySqlStmt.SqlLAB004Read(txtTestCd.Text, tmpSpc, tmpSeq)
'
'    lblSpcCd.Caption = tmpSpc    'tabSpecimen.SelectedItem.Name
'    lblSpcName.Caption = tabSpecimen.SelectedItem.Caption

    tmpSpc = Mid(tabSpecimen.SelectedItem.Key, 2)
    tmpSeq = CStr(tabSpecimen.SelectedItem.Index)
    tmpSql = MySqlStmt.SqlLAB004Read(txtTestCd.Text, tmpSpc, tmpSeq)

    lblSpcCd.Caption = tmpSpc    'tabSpecimen.SelectedItem.Name
    lblSpcName.Caption = tabSpecimen.SelectedItem.Caption

    Call Lab004Load(tmpSql)
    Call Lab004Show

End Sub

'% 검체를 선택하면 상세정보 Display


Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objCodeList Is Nothing Then Call cmdPopupList_Click
        'Call objCodeList.SetFocus(2)
    End If
End Sub

'% 검사코드를 입력 후 엔터키를 눌렀을 경우 데이타 조회....

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim tmpSql As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Not ConfirmExit Then
       KeyAscii = 0
       Exit Sub
    End If

    If KeyAscii = vbKeyReturn Then
       If txtTestCd.Text = "" Then Exit Sub
       Call ClearRtn(5)
       tmpSql = MySqlStmt.SqlLAB001Read(Trim(txtTestCd.Text))
       Call Lab001Load(tmpSql)
       Call Lab001Show
       tmpSql = MySqlStmt.SqlSpecimenRead(Trim(txtTestCd.Text))
       Call LabSpecimenLoad(tmpSql)
    
    
    
    End If

End Sub


'% Sub Routine 1 : Lab001Load
'%                        Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                        클래스 clsItem의 Data Attribute에 저장한다.

Function Lab001Load(ByVal SqlStmt As String)

    Dim MyRs As Recordset       'Oracle DynaSet

    On Error GoTo Error_Trap

    Set MyRs = New Recordset   'Sql 실행
    MyRs.Open SqlStmt, DBConn
    
    MyItems.Clear
    If MyRs.EOF Then GoTo NoData

    With MyItem
        While (MyRs.EOF = False)
            .TestCd = Trim("" & MyRs.Fields("TestCd").Value)
            .ApplyDt = Trim("" & MyRs.Fields("ApplyDt").Value)
            .TestNm = Trim("" & MyRs.Fields("TestNm").Value)
            .AbbrNm5 = Trim("" & MyRs.Fields("AbbrNm5").Value)
            .AbbrNm10 = Trim("" & MyRs.Fields("AbbrNm10").Value)
            .WorkArea = Trim("" & MyRs.Fields("WorkArea").Value)
            .RstType = Trim("" & MyRs.Fields("RstType").Value)
            .TestDiv = Trim("" & MyRs.Fields("TestDiv").Value)
            .RptSeq = Val("" & MyRs.Fields("RptSeq").Value)
            .PanelFg = Trim("" & MyRs.Fields("PanelFg").Value)
            .DetailFg = Trim("" & MyRs.Fields("DetailFg").Value)
            .TxtType = Trim("" & MyRs.Fields("TxtType").Value)
            .RstDiv = Trim("" & MyRs.Fields("RstDiv").Value)
            .OutLabCd = Trim("" & MyRs.Fields("OutLabCd").Value)
            .GrpFg = Trim("" & MyRs.Fields("GrpFg").Value)
            .WorkUnit = Val("" & MyRs.Fields("WorkUnit").Value)
            .AttrCd = Trim("" & MyRs.Fields("AttrCd").Value)
            .ExpDt = Trim("" & MyRs.Fields("ExpDt").Value)
            .ItemSeq = Val("" & MyRs.Fields("ItemSeq").Value)
            .GroupCd = Trim("" & MyRs.Fields("GroupCd").Value)
            MyItems.Add Trim("" & MyRs.Fields("ApplyDt").Value), MyItem
            MyRs.MoveNext
        Wend
    End With
    Lab001Load = True
    Exit Function

NoData:
    Set MyRs = Nothing
    Lab001Load = False
    Exit Function

Error_Trap:
    If Err.Number <> 94 Then
        MsgBox Err.Number & "  " & Err.Description
        Exit Function
    Else
        Resume Next
    End If

End Function


'% Sub Routine 2 : Lab001Move
'%                        각 화면상에 입력된 필드값을 Specimen 클래스로
'%                        치환한다.

Sub Lab001Move(ByRef MyItem As clsItem)

    Dim i As Integer

    With MyItem
        .TestCd = txtTestCd.Text
        .ApplyDt = Format(dtpApplyDate.Value, CS_DateDbFormat)
        .TestNm = txtFullNm.Text
        .AbbrNm5 = txtAbbrNm5.Text
        .AbbrNm10 = txtAbbrNm10.Text
        .WorkArea = medGetP(cboWorkArea.Text, 1, " ")
        .RstType = medGetP(cboRstType.Text, 1, " ")
        .RptSeq = Val(txtRptSeq.Text)
        .DetailFg = Choose(chkDetailFg.Value + 1, "", "*")
        .OutLabCd = Trim(medGetP(cboRefLab.Text, 1, " "))
        .GrpFg = chkGrpFg.Value
        .WorkUnit = Val(txtWorkUnit.Text)
        .AttrCd = medGetP(cboAttrCd.Text, 1, " ")

        .GroupCd = medGetP(cboStGroup.Text, 1, " ")
        .ItemSeq = Val(txtItemSeq.Text)

        If IsNull(dtpExpireDate.Value) Then
            .ExpDt = ""
        Else
            .ExpDt = Format(dtpExpireDate.Value, CS_DateDbFormat)
        End If

        '그룹처방여부
        If optPanFg(0).Value Then .PanelFg = ""
        If optPanFg(1).Value Then .PanelFg = "G"
        If optPanFg(2).Value Then .PanelFg = "D"

        '검사구분
        For i = 0 To optTestDiv.Count - 1
           If optTestDiv(i).Value Then .TestDiv = i
        Next

        'Text Type
        For i = 0 To optTextFg.Count - 1
           If optTextFg(i).Value Then .TxtType = i
        Next

        'Alternative / Required
        For i = 0 To optRstDiv.Count - 1
           If optRstDiv(i).Value Then .RstDiv = Choose(i + 1, "R", "A")
        Next

    End With

End Sub


'% Sub Routine 3 : LabSpecimenLoad
'%                        지정검체명들을 Tab에 Display

Sub LabSpecimenLoad(ByVal SqlStmt As String)

    Dim MyRs     As Recordset       'Oracle DynaSet
    Dim i        As Integer
    Dim tmpSpcCd As String
    Dim tmpSpcNm As String

    
    
    Set MyRs = New Recordset   'Sql 실행
    MyRs.Open SqlStmt, DBConn
    
    i = 0
    
    
    tabSpecimen.Tabs.Clear
    
    While (MyRs.EOF = False)
        tmpSpcCd = "" & MyRs.Fields("SpcCd").Value
        tmpSpcNm = "" & MyRs.Fields("SpcNm").Value
        
        
        tabSpecimen.Tabs.Add , "S" & Trim(tmpSpcCd), tmpSpcNm
        
        
        
        MyRs.MoveNext
        i = i + 1
    Wend
        
    'If tabSpecimen.Pages.Count > 0 Then tabSpecimen.Value = 0: Call tabSpecimen_Click(0)
    If tabSpecimen.Tabs.Count > 0 Then tabSpecimen.Tabs(1).Selected = True
    
    Set MyRs = Nothing

End Sub


'% Sub Routine 4 : Lab004Load
'%                        Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                        클래스 clsItem의 Data Attribute에 저장한다.

Sub Lab004Load(ByVal SqlStmt As String)

    Dim MySpecimen As clsSpecimen
    Dim MyRs As Recordset

    On Error GoTo Error_Trap

    Set MySpecimen = New clsSpecimen
    Set MyRs = New Recordset   'Sql 실행
    MyRs.Open SqlStmt, DBConn
    
    MySpecimens.Clear
    With MySpecimen
        While (MyRs.EOF = False)
            .TestCd = "" & MyRs.Fields("TestCd").Value
            .SpcCd = "" & MyRs.Fields("SpcCd").Value
            .Seq = Val("" & MyRs.Fields("Seq").Value)
            .ApplyDt = "" & MyRs.Fields("ApplyDt").Value
            .SpcGrpCd = "" & MyRs.Fields("SGroup").Value
            .LabelCnt = Val("" & MyRs.Fields("LabelCnt").Value)
            .RstUnit = "" & MyRs.Fields("RstUnit").Value
            .RndFg = "" & MyRs.Fields("RndFg").Value
            .StatFg = "" & MyRs.Fields("StatFg").Value
            .AvalVal = Val("" & MyRs.Fields("AvalVal").Value)
            .PanicFg = "" & MyRs.Fields("PanicFg").Value
            .PanicFrVal = Val("" & MyRs.Fields("PanicFrVal").Value)
            .PanicToVal = Val("" & MyRs.Fields("PanicToVal").Value)
            .DeltaFg = "" & MyRs.Fields("DeltaFg").Value
            .DeltaVal1 = Val("" & MyRs.Fields("DeltaVal").Value)
            .DeltaVal2 = Val("" & MyRs.Fields("DeltaVal2").Value)
            .TestCost = "" & MyRs.Fields("TestCost").Value
            .StoreCd = "" & MyRs.Fields("StoreCd").Value
            .TatAvg = Val("" & MyRs.Fields("TatAvg").Value)
            .SpcQty = Val("" & MyRs.Fields("SpcQty").Value)
            .SpcUnit = "" & MyRs.Fields("SpcUnit").Value
            .ExpDt = "" & MyRs.Fields("ExpDt").Value
            MySpecimens.Add MyRs.Fields("ApplyDt").Value, MySpecimen
            MyRs.MoveNext
        Wend
    End With

    'MyItem.RemoveParameters
    Set MyRs = Nothing
    Exit Sub

Error_Trap:
    If Err.Number <> 94 Then
        MsgBox Err.Number & "  " & Err.Description
        Set MyRs = Nothing
        Exit Sub
    Else
        Resume Next
    End If

End Sub


'% 검사항목 마스터의 적용일(ApplyDt)을 Tab에 Display

Sub Lab001Show()

    Dim i As Integer

'    tabItem.Pages.Clear
'    For i = 1 To MyItems.Count
'        tabItem.Pages.Add , Format(MyItems.Item(i).ApplyDt, CS_DateMask), i - 1
'    Next

    TabItem.Tabs.Clear
    For i = 1 To MyItems.Count
        TabItem.Tabs.Add i, , Format(MyItems.Item(i).ApplyDt, CS_DateMask)
    Next
    
'    If tabItem.Pages.Count > 0 Then
    If MyItems.Count > 0 Then
        TabItem.Tabs(1).Selected = True
        'tabItem.Value = 0
        'Call tabItem_Click
    Else
        cmdNew.Enabled = True
        InsertFlag = 0
        Call cmdNew_Click
    End If

End Sub


'% 검체 마스터의 검체(SpcCd)와 적용일(ApplyDt)을 Tab에 Display

Sub Lab004Show()

    Dim i As Integer

'    tabAppDt.Pages.Clear
'    For i = 1 To MySpecimens.Count
'        tabAppDt.Pages.Add , Format(MySpecimens.Specimen(i).ApplyDt, CS_DateMask), i - 1
'    Next
'    If tabAppDt.Pages.Count > 0 Then tabAppDt.Value = 0: Call tabAppDt_Click(0)
    tabAppDt.Tabs.Clear
    For i = 1 To MySpecimens.Count
        tabAppDt.Tabs.Add i, , Format(MySpecimens.Specimen(i).ApplyDt, CS_DateMask)
    Next
    
    If tabAppDt.Tabs.Count > 0 Then tabAppDt.Tabs(1).Selected = True

End Sub


Sub ClearRtn(ByVal intPart As Integer)

   ' intPart : 1-Item, 2-Clinical Notice, 3-Specimen, 4-Reference Range, 5-All

    Select Case intPart
        Case 1, 5: GoTo Clear1
        Case 2: GoTo Clear2
    End Select
    Exit Sub

Clear1:
        dtpExpireDate.Value = ""
        txtAbbrNm5.Text = ""
        txtAbbrNm10.Text = ""
        txtFullNm.Text = ""
        cboWorkArea.ListIndex = -1
        cboRstType.ListIndex = -1
        txtRptSeq.Text = "0"
        txtWorkUnit.Text = ""
        cboRefLab.ListIndex = -1
        optPanFg(0).Value = True
        optTestDiv(0).Value = True
        optRstDiv(0).Value = True
        optTextFg(0).Value = True
        cboAttrCd.ListIndex = -1
        chkDetailFg.Value = 0
        chkGrpFg.Value = 0
        cboStGroup.ListIndex = -1
        txtItemSeq.Text = "0"

        
        Call LockRtn(5, False)
        
        SvApplyDt = ""
        If intPart <> 5 Then Exit Sub

Clear2:
'        tabItem.Pages.Clear
'        tabSpecimen.Pages.Clear
'        tabAppDt.Pages.Clear
        TabItem.Tabs.Clear
        tabSpecimen.Tabs.Clear
        tabAppDt.Tabs.Clear
        lblSpcCd.Caption = ""
        lblSpcName.Caption = ""
        lblSpcAppDt.Caption = CS_BlankMask
        lblSpcExpDt.Caption = CS_BlankMask
        lblRstUnit.Caption = ""
        lblAvalVal.Caption = ""
        txtStoreCd.Text = ""
        lblSpcQty.Caption = ""
        lblSpcUnit.Caption = ""
        lblTestCost.Caption = ""
        lblTatAvg.Caption = ""
        chkRndFg.Value = 0
        chkStatFg.Value = 0
        chkPanicFg.Value = 0
        chkDeltaFg.Value = 0
        lblPanicFrVal.Caption = ""
        lblPanicToVal.Caption = ""
        lblDeltaVal1.Caption = ""
        lblDeltaVal2.Caption = ""
        lblLabelCnt.Caption = ""
        lblSpcGrpCd.Caption = ""
        If intPart <> 5 Then Exit Sub

Clear5:
        cmdEdit.Caption = "수정"
        cmdNew.Caption = "추가"
        cmdEdit.Enabled = True
        cmdNew.Enabled = True
        cmdCancel.Enabled = False
        UpdateFlag = 0
        InsertFlag = 0

End Sub

Sub LockRtn(ByVal intPart As Integer, ByVal LockValue As Boolean)

     Dim EnableValue As Boolean

    ' intPart : 1-Item, 2-Clinical Notice, 3-Specimen, 4-Reference Range, 5-All
    If LockValue Then
        EnableValue = False
    Else: EnableValue = True
    End If

    Select Case intPart
        Case 1, 5:
            If intPart = 1 Then dtpApplyDate.Enabled = EnableValue
            dtpExpireDate.Enabled = EnableValue
            txtAbbrNm5.Locked = LockValue
            txtAbbrNm10.Locked = LockValue
            txtFullNm.Locked = LockValue
            cboWorkArea.Locked = LockValue
            cboRstType.Locked = LockValue
            txtRptSeq.Locked = LockValue
            txtWorkUnit.Locked = LockValue
            cboRefLab.Locked = LockValue
            picPanelFg.Enabled = EnableValue
            picTestDiv.Enabled = EnableValue
            picRstDiv.Enabled = EnableValue
            picTextFg.Enabled = EnableValue
            chkGrpFg.Enabled = EnableValue
            UpDown1.Enabled = EnableValue
            UpDown2.Enabled = EnableValue
            txtItemSeq.Locked = LockValue
            cboStGroup.Locked = LockValue
            
    
        Case 2, 5:
        Case 3, 5:
        Case 4, 5:
    End Select

End Sub


Private Function ConfirmExit() As Boolean

    Dim intResp As VbMsgBoxResult

    ConfirmExit = True
    If InsertFlag = 1 Or UpdateFlag = 1 Then
        intResp = MsgBox("변경된 내용을 취소하시겠습니까 ? ", vbYesNo)
        If intResp = vbNo Then
            ConfirmExit = False
            Exit Function
        End If
        InsertFlag = 0
        UpdateFlag = 0
    End If

End Function

