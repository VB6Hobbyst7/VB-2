VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm101Attributes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "처방속성 등록"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "종류 (&X)"
      Height          =   405
      Left            =   8145
      TabIndex        =   40
      Top             =   5265
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "저장 (&S)"
      Height          =   405
      Left            =   7005
      TabIndex        =   39
      Top             =   5265
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabHeight       =   882
      BackColor       =   14737632
      ForeColor       =   5584725
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bone marrow biopsy / Leukocyte marker study"
      TabPicture(0)   =   "Lis1011.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Triple Marker"
      TabPicture(1)   =   "Lis1011.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "DTPicker3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "DTPicker2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text6"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "DTPicker1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text7"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Arterial blood gas"
      TabPicture(2)   =   "Lis1011.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text8"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text9"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Creatinine clearance"
      TabPicture(3)   =   "Lis1011.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label18"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label19"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label20"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label21"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Text10"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Text11"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Warning / Infection"
      TabPicture(4)   =   "Lis1011.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Check1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Check2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Check3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.CheckBox Check3 
         Caption         =   "알러지 ( A : Allergy )"
         Height          =   480
         Left            =   1035
         TabIndex        =   38
         Top             =   2670
         Width           =   2865
      End
      Begin VB.CheckBox Check2 
         Caption         =   "각종 유전병 ( H : hereditary )"
         Height          =   420
         Left            =   1035
         TabIndex        =   37
         Top             =   2175
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         Caption         =   "각종 전염병 ( I : Infection )"
         Height          =   300
         Left            =   1035
         TabIndex        =   36
         Top             =   1740
         Width           =   2655
      End
      Begin VB.TextBox Text11 
         Height          =   315
         Left            =   -72855
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   1980
         Width           =   1170
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   -72855
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   1545
         Width           =   1170
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   -72885
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2085
         Width           =   1170
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   -72885
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   1650
         Width           =   1170
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   -73065
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2925
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   -73065
         TabIndex        =   17
         Top             =   2460
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25296897
         CurrentDate     =   36843
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   -73065
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1995
         Width           =   1170
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   -73065
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox Text4 
         Height          =   660
         Left            =   -72825
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   8
         Top             =   4050
         Width           =   6645
      End
      Begin VB.TextBox Text3 
         Height          =   660
         Left            =   -72825
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   7
         Top             =   3195
         Width           =   6645
      End
      Begin VB.TextBox Text2 
         Height          =   660
         Left            =   -72810
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   6
         Top             =   2370
         Width           =   6645
      End
      Begin VB.TextBox Text1 
         Height          =   660
         Left            =   -72795
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   5
         Top             =   1545
         Width           =   6645
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   -68820
         TabIndex        =   18
         Top             =   2475
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25296897
         CurrentDate     =   36843
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   -68820
         TabIndex        =   19
         Top             =   2925
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25296897
         CurrentDate     =   36843
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "㎏"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71580
         TabIndex        =   35
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "㎝"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71580
         TabIndex        =   34
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 신장 :"
         Height          =   180
         Left            =   -74205
         TabIndex        =   31
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 체중 :"
         Height          =   180
         Left            =   -74205
         TabIndex        =   30
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "㏄"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71610
         TabIndex        =   29
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "℃"
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
         Left            =   -71610
         TabIndex        =   28
         Top             =   1710
         Width           =   180
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 체온 :"
         Height          =   180
         Left            =   -74235
         TabIndex        =   25
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ O2 량 :"
         Height          =   180
         Left            =   -74235
         TabIndex        =   24
         Top             =   2145
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "㎏"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71805
         TabIndex        =   23
         Top             =   2985
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "(회)"
         Height          =   180
         Left            =   -71790
         TabIndex        =   22
         Top             =   2070
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(주)"
         Height          =   180
         Left            =   -71790
         TabIndex        =   21
         Top             =   1620
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ Weight :"
         Height          =   180
         Left            =   -74415
         TabIndex        =   14
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 실제 생년월일 :"
         Height          =   180
         Left            =   -70500
         TabIndex        =   13
         Top             =   3015
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 임신 주수 :"
         Height          =   180
         Left            =   -74415
         TabIndex        =   12
         Top             =   1635
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 임신 횟수 :"
         Height          =   180
         Left            =   -74415
         TabIndex        =   11
         Top             =   2070
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ LMP :"
         Height          =   180
         Left            =   -74415
         TabIndex        =   10
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ 분만예정일 :"
         Height          =   180
         Left            =   -70500
         TabIndex        =   9
         Top             =   2565
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ CBC Result :"
         Height          =   180
         Left            =   -74640
         TabIndex        =   4
         Top             =   4065
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ Disagnosis :"
         Height          =   180
         Left            =   -74655
         TabIndex        =   3
         Top             =   3195
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ History :"
         Height          =   180
         Left            =   -74640
         TabIndex        =   2
         Top             =   2370
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "◈ Chief Complaint :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   1
         Top             =   1560
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frm101Attributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
