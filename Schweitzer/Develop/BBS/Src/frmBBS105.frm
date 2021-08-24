VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS105 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Out Patient Collection"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmBBS105.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   8460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS105.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS105.frx":0B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS105.frx":0EA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00F4F0F2&
      Caption         =   "채혈(&S)"
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
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15401"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
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
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
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
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   3675
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자 기본 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자검색"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8175
      Left            =   75
      TabIndex        =   5
      Top             =   285
      Width           =   3585
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Search"
         Height          =   630
         Left            =   60
         TabIndex        =   8
         Tag             =   "136"
         Top             =   525
         Width           =   3450
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&ID"
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
            Index           =   0
            Left            =   1680
            TabIndex        =   11
            TabStop         =   0   'False
            Tag             =   "15304"
            Top             =   285
            Width           =   495
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Name"
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
            Index           =   1
            Left            =   2205
            TabIndex        =   10
            TabStop         =   0   'False
            Tag             =   "15305"
            Top             =   270
            Width           =   810
         End
         Begin VB.TextBox txtSearchKey 
            Height          =   300
            Left            =   90
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   1470
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6855
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1185
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "환자ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "환자명"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "주민번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "생년월일"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "성별"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "접수일"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "접수번호"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "처방일자"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   330
         Left            =   1020
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   150
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
         Format          =   62455808
         UpDown          =   -1  'True
         CurrentDate     =   36342.5951388889
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   150
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
         Caption         =   "처방일"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   2130
      Left            =   3675
      TabIndex        =   6
      Top             =   285
      Width           =   10800
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   210
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   3450
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   210
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "성명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   3450
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   630
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   7005
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   7005
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   180
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   630
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "영수증번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   11235
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   525
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
         Caption         =   "병동"
         Appearance      =   0
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   17
         Top             =   210
         Width           =   2010
      End
      Begin VB.TextBox txtReceptNo 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   2010
      End
      Begin VB.TextBox txtRemark 
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
         Height          =   945
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1065
         Width           =   8880
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   4560
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   210
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   635
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
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   8115
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   630
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
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
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   360
         Left            =   8115
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   210
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
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
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   4560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
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
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   180
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "Remark"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   10
      Left            =   3675
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2415
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "처방 정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   5805
      Left            =   3675
      TabIndex        =   7
      Top             =   2655
      Width           =   10800
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "All Select(&A)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   30
         Tag             =   "137"
         Top             =   180
         Width           =   1530
      End
      Begin MSComctlLib.TabStrip tabCheckValue 
         Height          =   330
         Left            =   1710
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   582
         Style           =   2
         Separators      =   -1  'True
         ImageList       =   "imgList"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "접수(&1)"
               ImageVarType    =   2
               ImageIndex      =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "보류(&2)"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "제외(&3)"
               ImageVarType    =   2
               ImageIndex      =   3
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
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5115
         Left            =   165
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "10114"
         Top             =   525
         Width           =   10470
         _Version        =   196608
         _ExtentX        =   18468
         _ExtentY        =   9022
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
         GrayAreaBackColor=   15003117
         GridColor       =   14737632
         MaxCols         =   11
         OperationMode   =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS105.frx":123E
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   500
      End
   End
End
Attribute VB_Name = "frmBBS105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn
    tcSEL = 1
    tcORDDT
    tcORDNO
    tcTESTNM
    tcDOCTNM
    tcERCECHK
    tcTRANSDT
    tcTESTLOCATION
    tcRECEPTNO
    tcORDSEQ
    tcREMARK
End Enum
Private blnSearch As Boolean    'id/name 찾기 구분
Private strDeptCd As String     '진료과구분
Private strBlgCd As String      '병동의 건물 코드
Private strErbldcd As String    '응급일경우 검사할 건물코드
Private strGbldcd As String     '일반일경우 검사할 건물코드
Private blnAdd_Col As Boolean   '추가검체채혈(FALSE)과 일반 채혈(True)의 구분
Private strStatFg As String

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    lvwPtList.ListItems.Clear
    
    dtpToTime.value = Format(GetSystemDate, "yyyy-MM-dd")
    optSort(0).value = True
    blnSearch = True
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdClear_Click()
    Call ClearAll
End Sub

Private Sub ClearAll()
    lvwPtList.ListItems.Clear
    txtSearchKey = ""
    Clear
    
End Sub

Private Sub Clear()
    tblOrdSheet.MaxRows = 0: tblOrdSheet.MaxRows = 20
    'txtPtId = ""
    txtReceptNo = ""
    lblPtNm.Caption = ""
    lblOrdDt.Caption = ""
    lblSexAge.Caption = ""
    lblDeptNm.Caption = ""
    chkSelAll.value = 0
    Call ICSPatientMark
    
End Sub

Private Sub BarCode_Print(ObjDic As clsDictionary)
    Dim objBar As clsBarcode
    Dim strPtid As String
    Dim strPtNm As String
    Dim strSpcNo As String
    Dim strBuildNm As String        '건물이름
    Dim strW_Dept As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strAccSeq As String         'SpcYy-SpcNo 형태의 검체번호
    
    strW_Dept = strDeptCd
    Set objBar = New clsBarcode
    
'    Set objBar.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    strBuildNm = BBSName

    
    ObjDic.MoveFirst
    Do Until ObjDic.EOF
        strPtid = medGetP(ObjDic.GetString, 1, COL_DIV)
        strPtNm = medGetP(ObjDic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(ObjDic.GetString, 3, COL_DIV)
        strColDt = medGetP(ObjDic.GetString, 4, COL_DIV)
        strColTm = Mid(medGetP(ObjDic.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        
        
        '검체번호 출력 : 2001.2.8 추가
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        
        objBar.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                            strPtNm, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
                                            "", 1
        
        ObjDic.MoveNext
    Loop
    
    Set objBar = Nothing
    
End Sub

Private Sub cmdCollect_Click()
    Dim objCollect  As clsBBSCollection
    Dim ObjDic      As clsDictionary
    Dim objBar      As clsDictionary
    Dim strPtNm     As String       '환자명
    Dim strColID    As String      '채혈자
    Dim strColDt    As String      '채혈일
    Dim strColTm    As String      '채혈일시
    
    Dim i As Long
    
    If txtPtId = "" Then Exit Sub
    If Save_chk = False Then Exit Sub
    
    strPtNm = lblPtNm.Caption
    strColDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strColTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strColID = ObjSysInfo.EmpId
    
    Set objCollect = New clsBBSCollection
    Set ObjDic = New clsDictionary
    Set objBar = New clsDictionary
    
    ObjDic.Clear
    ObjDic.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd"
    
    ObjDic.AddNew txtPtId.Text, Join(Array(strPtNm, strColDt, strColTm, strColID, BBSBUSSDIV.stsNotBed, strBlgCd), COL_DIV)
    
    If objCollect.Set_Collect(ObjDic) Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
        '바코드 출력
            BarCode_Print objBar
        Else
            MsgBox "검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
        '환자리스트에서 삭제--------------------
        i = 0
        Do
            i = i + 1
            If i > lvwPtList.ListItems.Count Then Exit Sub

            If lvwPtList.ListItems(i).Text = txtPtId Then
                lvwPtList.ListItems.Remove i
                Exit Do
            End If
        Loop
        Call Clear
        txtPtId = ""
    End If
    Set objCollect = Nothing
    Set ObjDic = Nothing
    Set objBar = Nothing
    
    
End Sub
Private Function Save_chk() As Boolean
    Dim i As Integer
    
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcSEL
            If .value <> 0 Then
                Save_chk = True
                Exit For
            Else
                Save_chk = False
            End If
            
        Next i
    End With
    If Save_chk = False Then
        MsgBox "채혈대상을 선택한후 등록하십시오.", vbInformation + vbOKOnly, Me.Caption
    End If
End Function
'Private Sub lvwPtList_DblClick()
'    PtDisplay
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strRmk As String
    Dim i As Integer
    If Row < 1 Then Exit Sub
    
    With tblOrdSheet
        .Row = Row
        .Col = TblColumn.tcREMARK: txtRemark = .value
        If Col = TblColumn.tcSEL Then
            .Col = TblColumn.tcORDDT
            If .value <> "" Then
                .Col = TblColumn.tcSEL: .value = IIf(.value = 1, 0, 1)
                chkSelAll.value = 1
            End If
        End If
    End With
End Sub

Private Sub optSort_Click(Index As Integer)
    If Index = 0 Then
        blnSearch = True
    Else
        blnSearch = False
    End If
End Sub
Private Sub chkSelAll_Click()
    Dim i As Integer
    
    If chkSelAll.value = 1 Then
        With tblOrdSheet
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = TblColumn.tcORDDT
                If .value <> "" Then
                    .Col = TblColumn.tcSEL: .value = 1
                Else
                    Exit For
                End If
            Next
        End With
    Else
        With tblOrdSheet
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = TblColumn.tcORDDT
                If .value <> "" Then
                    .Col = TblColumn.tcSEL: .value = 1
                Else
                    Exit For
                End If
            Next
        End With
    End If
    
End Sub
Private Sub chkSelAll_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_GotFocus()
    txtPtId.SelStart = 0
    txtPtId.SelLength = Len(txtPtId)
    txtPtId.tag = txtPtId
End Sub

Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtPtId_LostFocus()
    Dim ii        As Long
    Dim strLng    As String
    
    If txtPtId.tag = txtPtId Then Exit Sub
    If Trim(txtPtId) = "" Then Exit Sub
    
    For ii = 1 To Val(BBS_PTID_LENGTH) - 1
        strLng = strLng & "0"
    Next ii
    

    If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
        txtPtId.Text = Format(txtPtId.Text, strLng & "#")
    End If
    
    
    If Direct_Collect(txtPtId.Text, True) = True Then
        Call lvwPtList_DblClick
        chkSelAll.SetFocus
    Else
        Call Clear
        txtPtId = ""
        MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요.", vbInformation + vbOKOnly, "채혈대상선택"
    End If
End Sub

Private Sub txtSearchKey_GotFocus()
    txtSearchKey.SelStart = 0
    txtSearchKey.SelLength = Len(txtSearchKey)
    txtSearchKey.tag = txtSearchKey
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub


Private Sub txtSearchKey_LostFocus()
    Dim ii        As Long
    Dim strLng    As String
    
    If Trim(txtSearchKey) = "" Then Exit Sub
    If txtSearchKey.tag = txtSearchKey Then Exit Sub
    
    If blnSearch = True Then
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii
        If Len(Trim(txtSearchKey.Text)) <> BBS_PTID_LENGTH Then
            txtSearchKey.Text = Format(txtSearchKey.Text, strLng & "#")
        End If
        If Direct_Collect(txtSearchKey.Text, blnSearch) = False Then
            MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요", vbInformation + vbOKOnly, "채혈대상선택"
        End If
    Else
        If Direct_Collect(txtSearchKey, blnSearch) = False Then
            MsgBox "조건에 맞는 자료가 없습니다." & vbCrLf & "확인후 조회하세요", vbInformation + vbOKOnly, "채혈대상선택"
        End If
    End If
    
    Call Clear
    txtSearchKey = ""
End Sub
Private Function Direct_Collect(searchkey As String, TF As Boolean) As Boolean
'채혈대상자를 조회시에 사용함.
'조회하고자 하는 문자를 입력한후 Enter(신규검체조회,추가검체 두가지를 구분하여 보여준다.
'처음 채혈하고자하는 채혈대상과 검체추가에 의한 채혈의 구분은
'리스트뷰 item 3,4,5 에 접수일자/접수번호를 가지고 구분한다.
    Dim objGetSql As New clsBBSCollection
    Dim DrRS As New Recordset
    Dim strOrdDt As String
    Dim blnEOF As Boolean
    Dim blnEOF1 As Boolean
    Dim iTmx As Object
    Dim itmx2 As Object
    Dim itmFound As ListItem
    
    strOrdDt = Format(dtpToTime.value, PRESENTDATE_FORMAT)
    
    lvwPtList.ListItems.Clear

    
    DrRS.Open objGetSql.Get_CollectOrder(searchkey, TF, BBSBUSSDIV.stsNotBed, strOrdDt), DBConn
    
    If DrRS.EOF = False Then
        With lvwPtList
            .ListItems.Clear
            Do Until DrRS.EOF
                Set itmx2 = .ListItems.Add(, , DrRS.Fields("ptid").value & "")
                itmx2.SubItems(1) = DrRS.Fields("ptnm").value & ""
                itmx2.SubItems(2) = Mid(DrRS.Fields("ssn").value & "", 1, 6) & "-" & _
                                    Mid(DrRS.Fields("ssn").value & "", 7)
                DrRS.MoveNext
            Loop
        End With
        blnEOF = True
    End If
    Set DrRS = Nothing
    
    '추가검체 조회

    Set DrRS = New Recordset
    DrRS.Open objGetSql.Get_AddSpcInFo(searchkey, TF), DBConn
    
    If DrRS.EOF = False Then
        With lvwPtList
            Do Until DrRS.EOF
                Set itmFound = .FindItem(DrRS.Fields("ptid").value & "", lvwText, , lvwPartial)
                If itmFound Is Nothing Then
                    Set iTmx = .ListItems.Add(, , DrRS.Fields("ptid").value & "")
                    iTmx.ForeColor = vbBlue
                    iTmx.SubItems(1) = DrRS.Fields("ptnm").value & ""
                    iTmx.ListSubItems(1).ForeColor = vbBlue
                    iTmx.SubItems(2) = Mid(DrRS.Fields("ssn").value & "", 1, 6) & "-" & _
                                       Mid(DrRS.Fields("ssn").value & "", 7)
                    iTmx.ListSubItems(2).ForeColor = vbBlue
                    iTmx.SubItems(3) = DrRS.Fields("accdt").value & ""
                    iTmx.SubItems(4) = DrRS.Fields("accno").value & ""
                    iTmx.SubItems(5) = DrRS.Fields("orddt").value & ""
                Else
                    '정상적인 채혈과 검체추가가 겹치는 경우
                    .ListItems(itmFound.Index).SubItems(3) = "*"
                    .ListItems(itmFound.Index).ForeColor = vbBlue
                    .ListItems(itmFound.Index).ListSubItems(1).ForeColor = vbBlue
                    .ListItems(itmFound.Index).ListSubItems(2).ForeColor = vbBlue
                    
                End If
                DrRS.MoveNext
            Loop
        End With
        blnEOF1 = True
    End If
    
    Set DrRS = Nothing
    
    If blnEOF = False And blnEOF1 = False Then
        Direct_Collect = False
        Clear
    Else
        Direct_Collect = True
    End If
    
    Set objGetSql = Nothing
End Function

Private Sub lvwPtList_DblClick()
    Dim iTmx As ListItem
    Dim strPtid As String
    Dim strAccDt As String
    Dim strAccSeq As String
    Dim tmpOrdDt As String
    
    With lvwPtList
        Set iTmx = .ListItems(.SelectedItem.Index)
        strPtid = .ListItems(.SelectedItem.Index).Text
        strAccDt = iTmx.SubItems(3)
        strAccSeq = iTmx.SubItems(4)
        tmpOrdDt = iTmx.SubItems(5)
    End With
    
    '감염관리
    Call ICSPatientMark(strPtid, enICSNum.BBS_ALL)
    
    If strAccDt = "" Or strAccDt = "*" Then         '처방에 따른 정상적인 채혈
        blnAdd_Col = True                           '(*)는 정상 채혈과 추가채혈이 동시존재한다.
        
        If strAccDt = "*" Then
            ptInfo strPtid, tmpOrdDt
        Else
            ptInfo strPtid, Format(dtpToTime.value, PRESENTDATE_FORMAT)
        End If
        PtDisplay strAccDt
        
    Else                                            '검체추가에따를 채혈
        blnAdd_Col = False
        ptInfo strPtid, , strAccDt, strAccSeq
        PtDisplay strAccDt, strAccSeq
    End If
End Sub
Private Sub ptInfo(ByVal PtId As String, Optional orddt As String = "", _
                   Optional accdt As String = "", Optional accseq As String = "")
'리스트뷰에서 선택한 환자의 환자정보와 채혈내역에 저장될 건물코드를 조회한다.
    Dim objGetSql As New clsGetSqlStatement
    Dim objCollect As New clsBBSCollection
    Dim DrRS      As New Recordset
    Dim strTmp    As String
    Dim strSDA    As String
    
    Set DrRS = objGetSql.Get_PtInfo(PtId, BBSBUSSDIV.stsNotBed, orddt, accdt, accseq)
    If DrRS.RecordCount > 0 Then
        txtPtId = PtId
        lblPtNm.Caption = DrRS.Fields("ptnm").value & ""
        
        strSDA = SDA_String(DrRS.Fields("ssn").value & "")
        lblSexAge.Caption = medGetP(strSDA, 1, COL_DIV) & "/" & medGetP(strSDA, 3, COL_DIV)
        
        lblDeptNm.Caption = IIf(IsNull(DrRS.Fields("deptnm").value & "") = True, "", DrRS.Fields("deptnm").value & "")
        strDeptCd = DrRS.Fields("deptcd").value & ""
        
        strBlgCd = ObjSysInfo.BuildingCd
        strTmp = objCollect.TestBuildCd(strBlgCd)
        strErbldcd = medGetP(strTmp, 1, COL_DIV)
        strGbldcd = medGetP(strTmp, 2, COL_DIV)
    End If
    
    Set objGetSql = Nothing
    Set objCollect = Nothing
End Sub
Private Sub PtDisplay(Optional ByVal accdt As String = "", Optional ByVal accseq As String = "")
    '조회된 환자ID를 가지고 채혈등록시 필요한 자료를 가지고 온다.
    Dim objGetSql As New clsBBSCollection
    Dim DrRS      As New Recordset
    Dim iTmx      As ListItem
    Dim strOrdDt  As String
    Dim strReqDt  As String
    Dim blnStatFg As Boolean
    Dim i As Integer
    
    

    
    strOrdDt = Format(dtpToTime.value, PRESENTDATE_FORMAT)
    i = 1
    If accdt = "" Then
        Set DrRS = objGetSql.Get_ORDER_105(txtPtId.Text, strOrdDt)
    Else
        Set DrRS = objGetSql.Get_ADDSPC(txtPtId.Text, accdt, accseq)
        If Not DrRS.EOF Then strReqDt = DrRS.Fields("reqdt1").value & ""
    End If
    
    With tblOrdSheet
        .ReDraw = False
        .MaxRows = 0: .MaxRows = 20
        Do Until DrRS.EOF = True
            .MaxRows = DrRS.RecordCount
            .Row = i
            strReqDt = DrRS.Fields("reqdt").value & ""
            .Col = TblColumn.tcORDDT:   .value = Mid(DrRS.Fields("orddt").value & "", 1, 4) & "-" & _
                                                 Mid(DrRS.Fields("orddt").value & "", 5, 2) & "-" & _
                                                 Mid(DrRS.Fields("orddt").value & "", 7)
            .Col = TblColumn.tcORDNO:   .value = Trim(DrRS.Fields("ordno").value & "")
            .Col = TblColumn.tcTESTNM:  .value = Get_TestNm(DrRS.Fields("ordcd").value & "")
            .Col = TblColumn.tcDOCTNM:  .value = GetDoctNm(DrRS.Fields("orddoct").value & "")
            .Col = TblColumn.tcERCECHK: .value = Trim(IIf(DrRS.Fields("statfg").value & "" = "1", "Y", ""))
            .ForeColor = RGB(255, 0, 0)
            .Col = TblColumn.tcTRANSDT: .value = Format(strReqDt, "####-##-##") & " " & _
                                                 Format(Mid(DrRS.Fields("reqtm").value & "", 1, 4), "00:00")
            Select Case DrRS.Fields("statfg").value & ""
                Case "1": .Col = TblColumn.tcTESTLOCATION: .value = objGetSql.TestBldNm(strErbldcd)
                Case "0": .Col = TblColumn.tcTESTLOCATION: .value = objGetSql.TestBldNm(strGbldcd)
            End Select
            
            .Col = TblColumn.tcRECEPTNO:   .value = DrRS.Fields("receptno").value & ""
            .Col = TblColumn.tcORDSEQ: .value = Trim(DrRS.Fields("ordseq").value & "")
            .Col = TblColumn.tcREMARK: .value = IIf(IsNull(DrRS.Fields("mesg").value & "") = True, "", DrRS.Fields("mesg").value & "")
            i = i + 1
            DrRS.MoveNext
        Loop
        For i = 1 To .MaxRows
            .Col = TblColumn.tcREMARK
            If .value <> "" Then
                txtRemark = txtRemark & .value & vbNewLine
            End If
            .Col = TblColumn.tcERCECHK
            If blnStatFg = False Then
                If .value = "Y" Then
                    strStatFg = "1"
                    blnStatFg = True
                End If
            End If
        Next
        If txtRemark <> "" Then
            txtRemark = Mid(txtRemark, 1, Len(txtRemark) - 1)
        End If
        .ReDraw = True
    End With
    Set DrRS = Nothing
    Set objGetSql = Nothing
End Sub







