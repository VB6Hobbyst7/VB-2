VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS811 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "수혈처방 마스터등록"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS811.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5430
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "128"
      Top             =   7860
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4080
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "124"
      Top             =   7860
      Width           =   1305
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   825
      Left            =   1140
      TabIndex        =   25
      Top             =   240
      Width           =   8430
      Begin VB.ListBox lstItemList 
         Height          =   240
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1605
         MaxLength       =   8
         TabIndex        =   1
         Top             =   300
         Width           =   1425
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
         Height          =   285
         Index           =   0
         Left            =   1155
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frmBBS811.frx":076A
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   285
         Width           =   300
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00CDE7FA&
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
         Height          =   420
         Index           =   1
         Left            =   6945
         Style           =   1  '그래픽
         TabIndex        =   15
         Tag             =   "124"
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00CDE7FA&
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
         Height          =   420
         Index           =   0
         Left            =   5685
         Style           =   1  '그래픽
         TabIndex        =   14
         Tag             =   "124"
         Top             =   225
         Width           =   1185
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
         Left            =   270
         TabIndex        =   17
         Tag             =   "35121"
         Top             =   345
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip tabItem 
      Height          =   390
      Left            =   1140
      TabIndex        =   26
      Top             =   1080
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   6285
      Left            =   1140
      TabIndex        =   16
      Top             =   1320
      Width           =   8445
      Begin VB.PictureBox picOrdDiv 
         BackColor       =   &H00DBE6E6&
         Height          =   360
         Left            =   2025
         ScaleHeight     =   300
         ScaleWidth      =   3630
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3180
         Width           =   3690
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Irradiation 처방"
            Height          =   300
            Index           =   1
            Left            =   1845
            TabIndex        =   47
            Tag             =   "35135"
            Top             =   15
            Width           =   1605
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "혈액 처방"
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   46
            Tag             =   "35136"
            Top             =   15
            Width           =   1140
         End
      End
      Begin VB.CheckBox chkNewTestDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "이코드는 혈액은행에서 출고한 내역으로 수가계산을 합니다."
         Height          =   180
         Left            =   600
         TabIndex        =   43
         Top             =   5820
         Width           =   5175
      End
      Begin VB.TextBox txtVolumn 
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
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   37
         Top             =   4740
         Width           =   2385
      End
      Begin VB.ComboBox cboRefLab 
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
         Height          =   315
         ItemData        =   "frmBBS811.frx":0CF4
         Left            =   2025
         List            =   "frmBBS811.frx":0CF6
         Style           =   2  '드롭다운 목록
         TabIndex        =   36
         Top             =   4410
         Width           =   3735
      End
      Begin VB.PictureBox picXMethod 
         BackColor       =   &H00DBE6E6&
         Height          =   360
         Left            =   2025
         ScaleHeight     =   300
         ScaleWidth      =   3630
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4020
         Width           =   3690
         Begin VB.OptionButton optXMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "None"
            Height          =   300
            Index           =   3
            Left            =   2760
            TabIndex        =   42
            Tag             =   "35136"
            Top             =   10
            Width           =   780
         End
         Begin VB.OptionButton optXMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Major"
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   35
            Tag             =   "35136"
            Top             =   10
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton optXMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Minor"
            Height          =   300
            Index           =   1
            Left            =   1005
            TabIndex        =   34
            Tag             =   "35135"
            Top             =   10
            Width           =   825
         End
         Begin VB.OptionButton optXMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Both"
            Height          =   300
            Index           =   2
            Left            =   1935
            TabIndex        =   33
            Tag             =   "35136"
            Top             =   10
            Width           =   720
         End
      End
      Begin VB.PictureBox picTestDiv 
         BackColor       =   &H00DBE6E6&
         Height          =   360
         Left            =   2025
         ScaleHeight     =   300
         ScaleWidth      =   3630
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3690
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "교환수혈"
            Height          =   300
            Index           =   2
            Left            =   2520
            TabIndex        =   44
            Tag             =   "35135"
            Top             =   15
            Width           =   1065
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "수혈처방"
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   31
            Tag             =   "35136"
            Top             =   15
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optTestDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Pheresis"
            Height          =   300
            Index           =   1
            Left            =   1305
            TabIndex        =   30
            Tag             =   "35135"
            Top             =   15
            Width           =   1065
         End
      End
      Begin VB.TextBox txtMatchCd 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   2010
         MaxLength       =   8
         TabIndex        =   7
         Top             =   5340
         Width           =   1425
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
         Index           =   1
         Left            =   3495
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "frmBBS811.frx":0CF8
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   5340
         Width           =   360
      End
      Begin VB.TextBox txtFullNm 
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
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   6
         Top             =   2460
         Width           =   3660
      End
      Begin VB.TextBox txtAbbrNm10 
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
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2100
         Width           =   2385
      End
      Begin VB.TextBox txtAbbrNm5 
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
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1740
         Width           =   1395
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "수정(&E)"
         Enabled         =   0   'False
         Height          =   405
         Left            =   3225
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   10
         Tag             =   "35106"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "추가(&A)"
         Height          =   405
         Left            =   2055
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   9
         Tag             =   "35106"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소(&U)"
         Height          =   405
         Left            =   4395
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '그래픽
         TabIndex        =   11
         Tag             =   "35106"
         Top             =   240
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtpApplyDate 
         Height          =   330
         Left            =   2070
         TabIndex        =   2
         Top             =   885
         Width           =   1785
         _ExtentX        =   3149
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62390275
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpExpireDate 
         Height          =   330
         Left            =   2070
         TabIndex        =   3
         Top             =   1260
         Width           =   1785
         _ExtentX        =   3149
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
         CheckBox        =   -1  'True
         CustomFormat    =   "yyy-MM-dd"
         DateIsNull      =   -1  'True
         Format          =   62390275
         CurrentDate     =   36328
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방구분"
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
         Left            =   660
         TabIndex        =   48
         Tag             =   "35120"
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Volume"
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
         Left            =   615
         TabIndex        =   41
         Tag             =   "35113"
         Top             =   4800
         Width           =   690
      End
      Begin VB.Label lblRefLab 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Component"
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
         Left            =   600
         TabIndex        =   40
         Tag             =   "35124"
         Top             =   4485
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사방법"
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
         Left            =   615
         TabIndex        =   39
         Tag             =   "35123"
         Top             =   4125
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "cc"
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
         Left            =   4440
         TabIndex        =   38
         Tag             =   "35113"
         Top             =   4860
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사구분"
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
         Left            =   660
         TabIndex        =   28
         Tag             =   "35120"
         Top             =   3660
         Width           =   780
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "수가코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label lblFullNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사명"
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
         Left            =   660
         TabIndex        =   23
         Tag             =   "35120"
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label lblAbbNm2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "약어2 (10문자)"
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
         Left            =   660
         TabIndex        =   22
         Tag             =   "35113"
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label lblAbbNm1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "약어1 (5문자)"
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
         Left            =   660
         TabIndex        =   21
         Tag             =   "35112"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblExpDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폐기일"
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
         Left            =   660
         TabIndex        =   20
         Tag             =   "35118"
         Top             =   1350
         Width           =   585
      End
      Begin VB.Label lblAppDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적용일"
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
         Left            =   660
         TabIndex        =   19
         Tag             =   "35114"
         Top             =   975
         Width           =   585
      End
      Begin VB.Label lblTestItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사항목 정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   360
         TabIndex        =   18
         Tag             =   "35131"
         Top             =   300
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmBBS811"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form명   : frmBBS805
'|  2. 기  능   : 수혈처방 마스터
'|  3. 작성자   : 김 동열
'|  4. 작성일   : 2000.11.20
'|
'|  CopyRight(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+
Option Explicit
Private objSql                  As clsBBSMSTStatement
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objListpop   As clsPopUpList
Attribute objListpop.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Dim RS As Recordset
    
    If txtTestCd = "" Then Exit Sub
    Set objSql = New clsBBSMSTStatement
    Set RS = objSql.LoadBBS001(txtTestCd.Text)
    If RS.EOF = False Then
        If cmdEdit.Enabled = False And cmdNew.Caption = "저장(&S)" Then
            cmdNew.Enabled = True
            cmdNew.Caption = "추가(&A)"
            cmdEdit.Enabled = True
            cmdCancel.Enabled = False
        ElseIf cmdNew.Enabled = False And cmdEdit.Caption = "저장(&S)" Then
            cmdNew.Enabled = True
            cmdEdit.Caption = "수정(&E)"
            cmdEdit.Enabled = True
            cmdCancel.Enabled = False
        End If
        TextLock
        txtTestCd.SetFocus
        Call tabItem_Click
    Else
        cmdNew.Enabled = True
        cmdNew.Caption = "추가(&A)"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = False
    End If
End Sub

Private Sub cmdClear_Click()
    Clear
    tabItem.Tabs.Clear
    txtTestCd.Text = ""
    TextLock
    txtTestCd.SetFocus
End Sub

Private Sub TextLock()
    dtpApplyDate.Enabled = False
    dtpExpireDate.Enabled = False
    txtAbbrNm5.Enabled = False
    txtAbbrNm10.Enabled = False
    txtFullNm.Enabled = False
    picOrdDiv.Enabled = False
    picTestDiv.Enabled = False
    picXMethod.Enabled = False
    cboRefLab.Enabled = False
    txtVolumn.Enabled = False
    txtMatchCd.Enabled = False
    cmdPopupList(1).Enabled = False
    chkNewTestDiv.Enabled = False
End Sub

Private Sub TextUnLock()
    dtpApplyDate.Enabled = True
    dtpExpireDate.Enabled = True
    txtAbbrNm5.Enabled = True
    txtAbbrNm10.Enabled = True
    txtFullNm.Enabled = True
    picOrdDiv.Enabled = True
    picTestDiv.Enabled = True
    picXMethod.Enabled = True
    cboRefLab.Enabled = True
    txtVolumn.Enabled = True
    txtMatchCd.Enabled = True
    cmdPopupList(1).Enabled = True
    chkNewTestDiv.Enabled = True
End Sub

Private Sub cmdEdit_Click()
    Dim RS As Recordset
    Dim rs1 As Recordset
    Dim strTmp As VbMsgBoxResult
    Dim strTmp1 As VbMsgBoxResult
    Dim strOpt As String
    Dim strCbo As String
    Dim strOrd As String
    
    
    Dim TDiv As String
    Dim Volume As Long
    Dim NewTestDiv As String
    
    '검사코드가 있는지 체크..
    If cmdEdit.Caption = "수정(&E)" Then
        cmdNew.Enabled = False
        cmdEdit.Caption = "저장(&S)"
        cmdCancel.Enabled = True
        
        Call TextUnLock
    Else
        If txtTestCd.Text = "" Then
            MsgBox "검사코드를 입력하여 주세요..", vbInformation, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        Set objSql = New clsBBSMSTStatement '등록여부체크..
        Set RS = objSql.GetBBS001(Trim(txtTestCd), Format(dtpApplyDate.Value, PRESENTDATE_FORMAT))
        If RS.EOF = False Then
            dtpApplyDate.Enabled = False
            dtpExpireDate.Enabled = True
            If IsNull(dtpExpireDate.Value) = True And Trim(RS.Fields("expdt").Value & "") = "" Then
                '수정여부체크..
                strTmp1 = MsgBox("수정하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
                If strTmp1 = vbCancel Then
                    Set RS = Nothing
                    Set objSql = Nothing
                    Clear
                    Exit Sub
                Else '수정
                    strCbo = medGetP(cboRefLab.Text, 1, " ")
                    If optXMethod(0).Value = True Then
                        strOpt = "0"
                    ElseIf optXMethod(1).Value = True Then
                        strOpt = "1"
                    ElseIf optXMethod(2).Value = True Then
                        strOpt = "2"
                    Else
                        strOpt = "3"
                    End If
                    
                    If optTestDiv(0).Value = True Then
                        TDiv = "0"
                        Volume = Val(Trim(txtVolumn))
                    ElseIf optTestDiv(1).Value = True Then
                        TDiv = "1"
                        Volume = 0
                    Else
                        TDiv = "2"
                        Volume = 0
                    End If
                    If chkNewTestDiv.Value = 1 Then
                        NewTestDiv = "Y"
                    ElseIf chkNewTestDiv.Value = 0 Then
                        NewTestDiv = "N"
                    Else
                        NewTestDiv = ""
                    End If
                    If optOrdDiv(0).Value = True Then
                        strOrd = "B"
                    Else
                        strOrd = "Z"
                    End If
                    
                    If objSql.InsertBBS001_Ghil(Trim(txtTestCd), Format(dtpApplyDate.Value, PRESENTDATE_FORMAT), Trim(txtFullNm), _
                                     Trim(txtAbbrNm5), Trim(txtAbbrNm10), Trim(strCbo), Volume, _
                                     TDiv, Trim(strOpt), Trim(txtMatchCd), NewTestDiv, False, strOrd) = True Then
                        MsgBox "수정하였습니다.", vbInformation, Me.Caption
                    End If
                End If
            Else
            '폐기여부체크..
                If Trim(RS.Fields("expdt").Value & "") = "" Then
                    If Format(dtpExpireDate, PRESENTDATE_FORMAT) < Format(GetSystemDate, PRESENTDATE_FORMAT) Then
                       MsgBox "이전날짜는 사용할 수 없습니다! 폐기일을 수정하세요..", vbInformation, Me.Caption
                       dtpExpireDate.SetFocus
                       Exit Sub
                    ElseIf Format(dtpExpireDate, PRESENTDATE_FORMAT) < RS.Fields("applydt").Value & "" Then
                       MsgBox "적용일 이전에 폐기할 수 없습니다! 폐기일을 수정하세요..", vbInformation, Me.Caption
                       dtpExpireDate.SetFocus
                       Exit Sub
                    End If
                    strTmp1 = MsgBox("폐기하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
                    If strTmp1 = vbCancel Then
                        Set RS = Nothing
                        Set objSql = Nothing
                        Clear
                        Exit Sub
                    Else '폐기
                        strCbo = medGetP(cboRefLab.Text, 1, " ")
                        If optXMethod(0).Value = True Then
                            strOpt = "0"
                        ElseIf optXMethod(1).Value = True Then
                            strOpt = "1"
                        ElseIf optXMethod(2).Value = True Then
                            strOpt = "2"
                        Else
                            strOpt = "3"
                        End If
                        If optTestDiv(0).Value = True Then
                            TDiv = "0"
                            Volume = Val(Trim(txtVolumn))
                        ElseIf optTestDiv(1).Value = True Then
                            TDiv = "1"
                            Volume = 0
                        Else
                            TDiv = "2"
                            Volume = 0
                        End If
                        If chkNewTestDiv.Value = 1 Then
                            NewTestDiv = "Y"
                        ElseIf chkNewTestDiv.Value = 0 Then
                            NewTestDiv = "N"
                        Else
                            NewTestDiv = ""
                        End If
                        If optOrdDiv(0).Value = True Then
                            strOrd = "B"
                        Else
                            strOrd = "Z"
                        End If

                        If objSql.InsertBBS001_Ghil(Trim(txtTestCd), Format(dtpApplyDate.Value, PRESENTDATE_FORMAT), Trim(txtFullNm), _
                                         Trim(txtAbbrNm5), Trim(txtAbbrNm10), Trim(strCbo), Volume, _
                                         TDiv, Trim(strOpt), Trim(txtMatchCd), NewTestDiv, False, strOrd, Format(dtpExpireDate.Value, PRESENTDATE_FORMAT)) = True Then
                            MsgBox "폐기하였습니다.", vbInformation, Me.Caption
                        End If
                    End If
                Else
                '재사용여부 체크..
                    strTmp1 = MsgBox("재사용하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
                    If strTmp1 = vbCancel Then
                        Set RS = Nothing
                        Set objSql = Nothing
                        Clear
                        Exit Sub
                    Else '재사용
                        strCbo = medGetP(cboRefLab.Text, 1, " ")
                        If optXMethod(0).Value = True Then
                            strOpt = "0"
                        ElseIf optXMethod(1).Value = True Then
                            strOpt = "1"
                        ElseIf optXMethod(2).Value = True Then
                            strOpt = "2"
                        Else
                            strOpt = "3"
                        End If
                        If optTestDiv(0).Value = True Then
                            TDiv = "0"
                            Volume = Val(Trim(txtVolumn))
                        ElseIf optTestDiv(1).Value = True Then
                            TDiv = "1"
                            Volume = 0
                        Else
                            TDiv = "2"
                            Volume = 0
                        End If
                        If chkNewTestDiv.Value = 1 Then
                            NewTestDiv = "Y"
                        ElseIf chkNewTestDiv.Value = 0 Then
                            NewTestDiv = "N"
                        Else
                            NewTestDiv = ""
                        End If
                        If optOrdDiv(0).Value = True Then
                            strOrd = "B"
                        Else
                            strOrd = "Z"
                        End If
                        
                        If objSql.InsertBBS001_Ghil(Trim(txtTestCd), Format(dtpApplyDate.Value, PRESENTDATE_FORMAT), Trim(txtFullNm), _
                                         Trim(txtAbbrNm5), Trim(txtAbbrNm10), Trim(strCbo), Volume, _
                                         TDiv, Trim(strOpt), Trim(txtMatchCd), NewTestDiv, False, strOrd) = True Then
                            MsgBox "재사용 할 수있습니다.", vbInformation, Me.Caption
                        End If
                    End If
                End If
            End If
        End If
        TextLock
        cmdEdit.Caption = "수정(&E)"
        cmdNew.Enabled = True
        cmdCancel.Enabled = False
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim i As Long
    Dim j As Long
    
    If txtTestCd.Text = "" Then Exit Sub
    Clear
    j = medListFind(lstItemList, txtTestCd.Text)
    If j < 0 Then Exit Sub
    Select Case Index
        Case 0:   'Previous
        If j <= 0 Then
            tabItemDisplay
            TextLock
            Exit Sub
        End If
        txtTestCd.Text = lstItemList.List(j - 1)
        Case 1:   'Next
        If j >= lstItemList.ListCount - 1 Then
            tabItemDisplay
            TextLock
            Exit Sub
        End If
        txtTestCd.Text = lstItemList.List(j + 1)
    End Select
    '내용 display..
    tabItemDisplay
    TextLock
End Sub

Private Sub cmdNew_Click()
    Dim RS As Recordset
    Dim strTmp As VbMsgBoxResult
    Dim strOpt As String
    Dim strCbo As String
    Dim strOrd As String
    
    '새로운 데이타 추가
    If cmdNew.Caption = "추가(&A)" Then
        If txtTestCd.Text = "" Then
            txtTestCd.SetFocus
            cmdNew.Caption = "추가(&A)"
            MsgBox "검사코드를 입력하여 주세요..", vbInformation, Me.Caption
            Exit Sub
        Else
            TextUnLock
            Clear
            cmdEdit.Enabled = False
            cmdNew.Caption = "저장(&S)"
            cmdNew.Enabled = True
            cmdCancel.Enabled = True
            dtpExpireDate.Enabled = False
            dtpApplyDate.SetFocus
        End If
        
    Else
        '저장여부 확인...
        Set objSql = New clsBBSMSTStatement
'        objSql.setDbConn DbConn
        If Format(dtpApplyDate, CS_DateDbFormat) < Format(GetSystemDate, CS_DateDbFormat) Then
            MsgBox "이전날짜는 사용할 수 없습니다! 적용일을 수정하세요..", vbInformation, Me.Caption
            dtpApplyDate.SetFocus
            Exit Sub
        End If
        strTmp = MsgBox("저장하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
        If strTmp = vbCancel Then
            Clear
            Set objSql = Nothing
            Exit Sub
        Else '저장
            Dim TDiv As String
            Dim Volume As Long
            Dim NewTestDiv As String
            
            strCbo = medGetP(cboRefLab.Text, 1, " ")
            If optXMethod(0).Value = True Then
                strOpt = "0"
            ElseIf optXMethod(1).Value = True Then
                strOpt = "1"
            ElseIf optXMethod(2).Value = True Then
                strOpt = "2"
            Else
                strOpt = "3"
            End If
            If optTestDiv(0).Value = True Then
                TDiv = "0"
                Volume = Val(Trim(txtVolumn))
            ElseIf optTestDiv(1).Value = True Then
                TDiv = "1"
                Volume = 0
            Else
                TDiv = "2"
                Volume = 0
            End If
            If optOrdDiv(0).Value = True Then
                strOrd = "B"
            Else
                strOrd = "Z"
            End If
            
            If chkNewTestDiv.Value = 1 Then
                NewTestDiv = "Y"
            ElseIf chkNewTestDiv.Value = 0 Then
                NewTestDiv = "N"
            Else
                NewTestDiv = ""
            End If
            If objSql.InsertBBS001_Ghil(Trim(txtTestCd), Format(dtpApplyDate, PRESENTDATE_FORMAT), Trim(txtFullNm), _
                                 Trim(txtAbbrNm5), Trim(txtAbbrNm10), Trim(strCbo), Volume, _
                                 TDiv, Trim(strOpt), Trim(txtMatchCd), NewTestDiv, True, strOrd) Then
                MsgBox "저장성공하였습니다.", vbInformation, Me.Caption
                tabItemDisplay
                List
            End If
        End If
        Clear
        cmdNew.Caption = "추가(&A)"
        cmdEdit.Caption = "수정(&E)"
        cmdEdit.Enabled = True
        Set RS = Nothing
        Set objSql = Nothing
    End If
End Sub

Private Sub cmdPopupList_Click(Index As Integer)
    '리스트 팝업을 불러오자...
    Set objSql = New clsBBSMSTStatement
    Set objListpop = New clsPopUpList
    objListpop.Connection = DBConn
'    objListpop.BackColor = Me.BackColor
    Select Case Index
        Case 0:
            objListpop.Tag = "TestCd"
            objListpop.FormCaption = "검사코드 찾기"
        Case 1:
            objListpop.Tag = "MatchCd"
            objListpop.FormCaption = "수가코드 찾기"
    End Select

    If Index = 0 Then
        Call objListpop.LoadPopup(objSql.GetPopup(Index)) ', 3300, 6300)
    Else
        '이부분 길병원 수가테이블이랑 매치시켜줘야 되거든요....
        Call objListpop.LoadPopup(objSql.GetPopup(0)) ', 5000, 9000)
    End If
        
    
    Set objSql = Nothing
End Sub

Private Sub dtpApplyDate_LostFocus()
    Dim RS As Recordset
    Dim rs1 As Recordset
    Dim strCompo As String
    Dim i As Long
    
    '저장되어 있는지 체크..
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.GetBBS001(txtTestCd, Format(dtpApplyDate.Value, PRESENTDATE_FORMAT))
    If RS.EOF = False Then
        TextLock
        dtpApplyDate.Value = Format(RS.Fields("applydt").Value & "", "##-##-##")
        txtAbbrNm5 = RS.Fields("abbrnm5").Value & ""
        txtAbbrNm10 = RS.Fields("abbrnm10").Value & ""
        txtFullNm = RS.Fields("testnm").Value & ""
        txtVolumn = RS.Fields("volumn").Value & ""
        txtMatchCd = RS.Fields("matchcd").Value & ""
        Select Case RS.Fields("xmethod").Value & ""
            Case "0": optXMethod(0).Value = True
            Case "1": optXMethod(1).Value = True
            Case "2": optXMethod(2).Value = True
            Case "3": optXMethod(3).Value = True
        End Select
        If Trim(RS.Fields("expdt").Value & "") = "" Then
            dtpExpireDate.Enabled = True
            dtpExpireDate.Value = ""
        Else
            dtpExpireDate.Enabled = True
            dtpExpireDate.Value = Format(RS.Fields("expdt").Value & "", "##-##-##")
        End If
        strCompo = RS.Fields("compocd").Value & ""
        cboRefLab.ListIndex = -1
        For i = 0 To cboRefLab.ListCount - 1
            If strCompo = medGetP(cboRefLab.List(i), 1, " ") Then
                cboRefLab.ListIndex = i
                Exit For
            End If
        Next i
        cmdNew.Enabled = True
        cmdNew.Caption = "추가(&A)"
        cmdEdit.Enabled = True
        cmdEdit.Caption = "수정(&E)"
        cmdCancel.Enabled = False
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
    txtTestCd.SetFocus
    List
End Sub

Private Sub List()
    Dim RS As Recordset
    Dim strTest As String
    
    
    '리스트박스에 검사코드를 넣어두자..
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.LoadAllBBS001
    If RS.EOF = False Then
        With lstItemList
            .Clear
            Do Until RS.EOF
                If strTest = RS.Fields("testcd").Value & "" Then
                    strTest = RS.Fields("testcd").Value & ""
                Else
                    lstItemList.AddItem RS.Fields("testcd").Value & ""
                    strTest = RS.Fields("testcd").Value & ""
                End If
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub
Private Sub Form_Load()
    Clear
    cboShow
End Sub

Private Sub objListpop_SendCode(ByVal SelString As String)
    '리스트박스에 있는내용을 가져오자..
    Select Case objListpop.Tag
    Case "TestCd"
        txtTestCd.Text = medGetP(SelString, 1, ";")
        tabItemDisplay
    Case "MatchCd"
        txtMatchCd.Text = medGetP(SelString, 1, ";")
    End Select
    Set objMyList = Nothing
    Set objListpop = Nothing
End Sub
Private Sub tabItemDisplay()
    Dim RS As Recordset
    Dim rs1 As Recordset
    Dim strCompo As String
    Dim i As Long
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.LoadBBS001(txtTestCd.Text)
    If RS.EOF = False Then
        dtpApplyDate.Enabled = False
        cmdNew.Enabled = True
        cmdNew.Caption = "추가(&A)"
        cmdEdit.Caption = "수정(&E)"
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        tabItem.Tabs.Clear
        Do Until RS.EOF = True
            i = i + 1
            tabItem.Tabs.Add i, , Format(RS.Fields("applydt").Value & "", "##-##-##")
            RS.MoveNext
        Loop
        Set RS = objSql.GetBBS001(txtTestCd, Format(tabItem.SelectedItem.Caption, PRESENTDATE_FORMAT))
        If RS.EOF = False Then
            dtpApplyDate.Enabled = False
            cmdNew.Enabled = True
            cmdNew.Caption = "추가(&A)"
            cmdEdit.Caption = "수정(&E)"
            cmdEdit.Enabled = True
            cmdCancel.Enabled = False
            TextLock
            dtpApplyDate.Value = Format(RS.Fields("applydt").Value & "", "##-##-##")
            txtAbbrNm5 = RS.Fields("abbrnm5").Value & ""
            txtAbbrNm10 = RS.Fields("abbrnm10").Value & ""
            txtFullNm = RS.Fields("testnm").Value & ""
            Select Case RS.Fields("testdiv").Value & ""
                Case "0": optTestDiv(0).Value = True
                Case "1": optTestDiv(1).Value = True
                Case "2": optTestDiv(2).Value = True
                Case Else:
                          optTestDiv(0).Value = False
                          optTestDiv(1).Value = False
                          optTestDiv(2).Value = False
            End Select
            Select Case RS.Fields("orddiv").Value & ""
                Case "B": optOrdDiv(0).Value = True
                Case "Z": optOrdDiv(1).Value = True
                Case Else:
                          optOrdDiv(0).Value = False
                          optOrdDiv(1).Value = False
            End Select
            txtVolumn = RS.Fields("volumn").Value & ""
            txtMatchCd = RS.Fields("matchcd").Value & ""
            Select Case RS.Fields("xmethod").Value & ""
                Case "0": optXMethod(0).Value = True
                Case "1": optXMethod(1).Value = True
                Case "2": optXMethod(2).Value = True
                Case "3": optXMethod(3).Value = True
                Case Else:
                          optXMethod(0).Value = False
                          optXMethod(1).Value = False
                          optXMethod(2).Value = False
                          optXMethod(3).Value = False
            End Select
            If Trim(RS.Fields("expdt").Value & "") = "" Or (Val(RS.Fields("expdt").Value & "") = 0) Then
                dtpExpireDate.Enabled = True
                dtpExpireDate.Value = ""
                dtpExpireDate.Enabled = False
            Else
                dtpExpireDate.Enabled = True
                
                dtpExpireDate.Value = Format(RS.Fields("expdt").Value & "", "##-##-##")
                dtpExpireDate.Enabled = False
            End If
            strCompo = RS.Fields("compocd").Value & ""
            cboRefLab.ListIndex = -1
            For i = 0 To cboRefLab.ListCount - 1
                If strCompo = medGetP(cboRefLab.List(i), 1, " ") Then
                    cboRefLab.ListIndex = i
                    Exit For
                End If
            Next i
            Select Case RS.Fields("newtestdiv").Value & ""
                Case "Y"
                    chkNewTestDiv.Value = 1
                Case "N"
                    chkNewTestDiv.Value = 0
                Case Else
                    chkNewTestDiv.Value = 2
            End Select
        End If
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub objListpop_SelectedItem(ByVal pSelectedItem As String)
    Select Case objListpop.Tag
    Case "TestCd"
        txtTestCd.Text = medGetP(pSelectedItem, 1, ";")
        tabItemDisplay
    Case "MatchCd"
        txtMatchCd.Text = medGetP(pSelectedItem, 1, ";")
    End Select
    Set objMyList = Nothing
    Set objListpop = Nothing
End Sub

Private Sub optTestDiv_Click(Index As Integer)
    Select Case Index
        Case 0
            txtVolumn.Enabled = True
        Case 1
            txtVolumn.Enabled = False
        Case 2
            txtVolumn.Enabled = False
    End Select
End Sub

'
Private Sub tabItem_Click()
    Dim RS As Recordset
    Dim rs1 As Recordset
    Dim strCompo As String
    Dim i As Long
    
    '선택된 적용일 따라 내용 display..
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.GetBBS001(txtTestCd, Format(tabItem.SelectedItem.Caption, PRESENTDATE_FORMAT))
    If RS.EOF = False Then
        dtpApplyDate.Enabled = False
        cmdNew.Enabled = True
        cmdNew.Caption = "추가(&A)"
        cmdEdit.Caption = "수정(&E)"
        cmdEdit.Enabled = True
        cmdCancel.Enabled = False
        dtpApplyDate.Value = Format(RS.Fields("applydt").Value & "", "##-##-##")
        TextLock
        txtAbbrNm5 = RS.Fields("abbrnm5").Value & ""
        txtAbbrNm10 = RS.Fields("abbrnm10").Value & ""
        txtFullNm = RS.Fields("testnm").Value & ""
        txtVolumn = RS.Fields("volumn").Value & ""
        txtMatchCd = RS.Fields("matchcd").Value & ""
        Select Case RS.Fields("xmethod").Value & ""
            Case "0": optXMethod(0).Value = True
            Case "1": optXMethod(1).Value = True
            Case "2": optXMethod(2).Value = True
            Case "3": optXMethod(3).Value = True
        End Select
        If Trim(RS.Fields("expdt").Value & "") = "" Then
            dtpExpireDate.Enabled = True
            dtpExpireDate.Value = ""
            dtpExpireDate.Enabled = False
        Else
            dtpExpireDate.Enabled = True
            dtpExpireDate.Value = Format(RS.Fields("expdt").Value & "", "##-##-##")
            dtpExpireDate.Enabled = False
        End If
        strCompo = RS.Fields("compocd").Value & ""
        cboRefLab.ListIndex = -1
        For i = 0 To cboRefLab.ListCount - 1
            If strCompo = medGetP(cboRefLab.List(i), 1, " ") Then
                cboRefLab.ListIndex = i
                Exit For
            End If
        Next i
        Select Case RS.Fields("newtestdiv").Value & ""
            Case "Y"
                chkNewTestDiv.Value = 1
            Case "N"
                chkNewTestDiv.Value = 0
            Case Else
                chkNewTestDiv.Value = 2
        End Select
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Clear()
    '깨끗이..
    cmdNew.Enabled = True
    cmdNew.Caption = "추가(&A)"
    cmdEdit.Enabled = False
    cmdCancel.Enabled = False
    dtpApplyDate.Enabled = True
    dtpApplyDate.Value = Format(Now, "YYYY-MM-DD")
    dtpExpireDate.Value = ""
    dtpExpireDate.Enabled = False
    txtAbbrNm5.Text = ""
    txtAbbrNm10.Text = ""
    txtFullNm.Text = ""
    txtVolumn.Text = ""
    txtMatchCd.Text = ""
    chkNewTestDiv.Value = 0
    cboRefLab.ListIndex = -1
End Sub

Private Sub txtAbbrNm10_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAbbrNm5_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtFullNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTestCd_LostFocus()
    Dim RS As Recordset
    
    If txtTestCd = "" Then Exit Sub
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DbConn
    Set RS = objSql.LoadBBS001(txtTestCd.Text)
    If RS.EOF = False Then
        tabItemDisplay
    Else
        Clear
        dtpApplyDate.Enabled = False
        dtpExpireDate.Enabled = False
        tabItem.Tabs.Clear
        TextLock
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub txtVolumn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboShow()
    Dim objCompo As clsComponent
    Dim RS As Recordset
    Dim strNm As String
    Dim i As Long
    
    Set objCompo = New clsComponent
    Set RS = objCompo.GetList(True)
    Set objCompo = Nothing
    
    If RS.EOF = False And RS.BOF = False Then
        With cboRefLab
            i = 0
            Do Until RS.EOF = True
                strNm = RS.Fields("compocd").Value & "" & " " & RS.Fields("componm").Value & ""
                .AddItem strNm, i
                ' 각 항목을 목록에 추가합니다.
                .Text = .List(0)
                i = i + 1
                RS.MoveNext
            Loop
        End With
        
        If cboRefLab.ListCount > 0 Then cboRefLab.ListIndex = -1
    End If
    Set RS = Nothing
End Sub



