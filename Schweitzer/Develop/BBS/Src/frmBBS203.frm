VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS203 
   BackColor       =   &H00DBE6E6&
   Caption         =   "보관검체관리"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmBBS203.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdAllExp 
      BackColor       =   &H00F4F0F2&
      Caption         =   "일괄폐기"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   27
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame DrFrame3 
      Height          =   495
      Left            =   75
      TabIndex        =   20
      Top             =   8550
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      Title           =   "Title"
      TitlePos        =   0
      DelLine         =   0
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
      Begin VB.ComboBox cboCol 
         Height          =   300
         Left            =   2700
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   90
         Width           =   735
      End
      Begin VB.ComboBox cboRow 
         Height          =   300
         Left            =   1410
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   90
         Width           =   735
      End
      Begin VB.ComboBox cboLeg 
         Height          =   300
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   90
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Col"
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
         Left            =   3495
         TabIndex        =   26
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Row"
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
         Left            =   2235
         TabIndex        =   25
         Top             =   150
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack"
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
         Left            =   885
         TabIndex        =   24
         Top             =   150
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdMove 
      Appearance      =   0  '평면
      BackColor       =   &H00F4F0F2&
      Caption         =   "보관장소이동"
      Height          =   510
      Left            =   4500
      Style           =   1  '그래픽
      TabIndex        =   19
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExp2 
      BackColor       =   &H00F4F0F2&
      Caption         =   "폐기후 보관"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExp 
      BackColor       =   &H00F4F0F2&
      Caption         =   "폐기"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame DrFrame1 
      Height          =   1185
      Left            =   75
      TabIndex        =   2
      Top             =   7260
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   2090
      Title           =   ""
      TitlePos        =   1
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   11565
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
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
         Caption         =   "병원명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   8730
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   120
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
         Caption         =   "환자명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   11565
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   510
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
         Caption         =   "경과시간"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   135
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   120
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
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   5865
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   510
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
         Caption         =   "접수자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   5865
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   8730
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   510
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
         Index           =   7
         Left            =   3000
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   510
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
         Caption         =   "채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNo 
         Height          =   360
         Left            =   1200
         TabIndex        =   7
         Top             =   120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblColDtTm 
         Height          =   360
         Left            =   4065
         TabIndex        =   8
         Top             =   120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblColNm 
         Height          =   360
         Left            =   4065
         TabIndex        =   9
         Top             =   510
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblRcvDtTm 
         Height          =   360
         Left            =   6945
         TabIndex        =   10
         Top             =   120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblRcvNm 
         Height          =   360
         Left            =   6945
         TabIndex        =   11
         Top             =   510
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
         Height          =   360
         Left            =   9795
         TabIndex        =   12
         Top             =   120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblPtId 
         Height          =   360
         Left            =   9795
         TabIndex        =   13
         Top             =   510
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblLocalNm 
         Height          =   360
         Left            =   12630
         TabIndex        =   14
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblPassTime 
         Height          =   360
         Left            =   12630
         TabIndex        =   15
         Top             =   510
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BackColor       =   14411494
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
   End
   Begin FPSpread.vaSpread tblLeg 
      Height          =   5985
      Left            =   75
      TabIndex        =   1
      Top             =   930
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   10557
      _StockProps     =   64
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridColor       =   14411494
      MaxCols         =   20
      MaxRows         =   10
      SelectBlockOptions=   0
      ShadowColor     =   14411494
      SpreadDesigner  =   "frmBBS203.frx":076A
   End
   Begin MSComctlLib.TabStrip tabLeg 
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LEG (A)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "LEG (B)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   555
      Left            =   75
      TabIndex        =   3
      Top             =   405
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   979
      Title           =   "LEG 상세 정보"
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MedControls1.LisLabel lblRowCnt 
         Height          =   360
         Left            =   1155
         TabIndex        =   4
         Top             =   90
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblColCnt 
         Height          =   360
         Left            =   3000
         TabIndex        =   5
         Top             =   90
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel lblRmk 
         Height          =   360
         Left            =   4845
         TabIndex        =   6
         Top             =   90
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   635
         BackColor       =   14411494
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   90
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
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
         Caption         =   "Row"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   1935
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   90
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
         Caption         =   "Col"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   3780
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   90
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
         Caption         =   "비고"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   75
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6930
      Width           =   14370
      _ExtentX        =   25347
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
      Caption         =   "  수혈 처방 리스트"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private BasicRowHeight As Double
Private BasicColWidth As Double
Private TableHeight As Double
Private TableWidth As Double

Private CurrentLegCd As String


Private Sub cboCol_Click()
'rack에 검체가 보관되지 않은 장소만 리스트에 보여줌.
'    Dim strSql As String
'    Dim Rs  As Recordset
'
'    strSql = " select * from " & T_BBS206 & _
'             " where " & DBW("centercd=", ObjSysInfo.BuildingCd) & _
'             " and " & DBW("legcd=", cboLeg.Text) & _
'             " and " & DBW("colno=", cboCol.Text)
'    Set Rs = New Recordset
'    Rs.Open strSql, DBConn
'
'    cboRow.Clear
'    Do Until Rs.EOF
'        If Rs.Fields("stscd").value & "" <> "1" Then
'            If medComboFind(cboRow, Rs.Fields("rowno").value & "") < 0 Then
'                cboRow.AddItem Rs.Fields("rowno").value & ""
'            End If
'
''            If medComboFind(cboCol, Rs.Fields("colno").value & "") < 0 Then
''                cboCol.AddItem Rs.Fields("colno").value & ""
''            End If
'        End If
'
'        Rs.MoveNext
'    Loop
'
'    Set Rs = Nothing
End Sub

Private Sub cboLeg_Click()
    Dim i As Long
    Dim RowCnt As Long
    Dim ColCnt As Long

    Dim objSpcM As clsSpecManagement

    Set objSpcM = New clsSpecManagement
    If objSpcM.ReadLegInfo(ObjSysInfo.BuildingCd, cboLeg.Text) = True Then
        RowCnt = objSpcM.RowCnt
        ColCnt = objSpcM.ColCnt
    End If
    Set objSpcM = Nothing

    cboRow.Clear
    For i = 1 To RowCnt
        cboRow.AddItem i
    Next i

    cboCol.Clear
    For i = 1 To ColCnt
        cboCol.AddItem i
    Next i
End Sub

Private Sub cboRow_Click()
'rack에 검체가 보관되지 않은 장소만 리스트에 보여줌.
    Dim strSql As String
    Dim Rs  As Recordset
    
    strSql = " select * from " & T_BBS206 & _
             " where " & DBW("centercd=", ObjSysInfo.BuildingCd) & _
             " and " & DBW("legcd=", cboLeg.Text) & _
             " and " & DBW("rowno=", cboRow.Text)
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    cboCol.Clear
    Do Until Rs.EOF
        If Rs.Fields("stscd").value & "" <> "1" Then
'            If medComboFind(cboRow, Rs.Fields("rowno").value & "") < 0 Then
'                cboRow.AddItem Rs.Fields("rowno").value & ""
'            End If
            
            If medComboFind(cboCol, Rs.Fields("colno").value & "") < 0 Then
                cboCol.AddItem Rs.Fields("colno").value & ""
            End If
        End If
        
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
End Sub

Private Sub cmdAllExp_Click()
'일괄폐기.....
    Dim objSpec As New clsSpecManagement
    Dim objPrgBar  As New clsProgress
    
    Dim LegCd As String
    Dim RowNo As String
    Dim ColNo As String
    
    Dim spcyy As String
    Dim spcno As String
    
    Dim SSQL  As String
    
    Dim ii    As Integer
    Dim jj    As Integer
    Dim Cnt   As Integer
    
    LegCd = tabLeg.SelectedItem.Key
    
'    Set objPrgBar.MyForm = Me
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = MainFrm.stsBar
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    With tblLeg
        objPrgBar.Max = .MaxRows
        For ii = 1 To .MaxRows
            .Row = ii
            RowNo = ii
            
            For jj = 1 To .MaxCols
                .Col = jj: ColNo = jj
                If .ForeColor = RGB(255, 0, 0) Then
                    spcyy = medGetP(.value, 1, "-")
                    spcno = medGetP(.value, 2, "-")

                    SSQL = objSpec.SetALL_Exp(ObjSysInfo.BuildingCd, LegCd, RowNo, ColNo)
                    DBConn.Execute SSQL
                    SSQL = objSpec.SetSpcExpString(spcyy, spcno)
                    DBConn.Execute SSQL
'                    If  = "01" Then
'                        sSql = objSpec.SetStorPositionUpdate_203(spcyy, spcno)
'                        DBConn.Execute sSql
'                    End If
                    Cnt = Cnt + 1
                End If
            Next jj
            objPrgBar.value = ii
        Next ii
    End With
    DBConn.CommitTrans
    MsgBox Cnt & "건의 검체가 폐기처리 되었습니다.", vbInformation + vbOKOnly, "검체일괄폐기"
    
    'Call tabLeg_Click
    
    Set objPrgBar = Nothing
    Set objSpec = Nothing
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    Set objPrgBar = Nothing
    Set objSpec = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExp_Click()
    '폐기
    Dim spcyy As String
    Dim spcno As String
    Dim value As String
    
    With tblLeg
        .Row = .ActiveRow
        .Col = .ActiveCol
        value = .value
        
        'value = Replace(value, vbNewLine, "-")
        spcyy = medGetP(value, 1, "-")
        spcno = medGetP(value, 2, "-")

        If Expire(spcyy, spcno, "1") = True Then
            Call QueryLegInfo(CurrentLegCd)
            Call ReadSpecList(CurrentLegCd)
            Call TblLegLeaveCell(.Col, .Row)
        End If
        
        .SetFocus
    End With
    
End Sub

Private Sub cmdExp2_Click()
    '폐기후 보관
    Dim spcyy As String
    Dim spcno As String
    Dim value As String
    
    With tblLeg
        .Row = .ActiveRow
        .Col = .ActiveCol
        value = .value
        
        'value = Replace(value, vbNewLine, "-")
        spcyy = medGetP(value, 1, "-")
        spcno = medGetP(value, 2, "-")

        If Expire(spcyy, spcno, "2") = True Then
            Call QueryLegInfo(CurrentLegCd)
            Call ReadSpecList(CurrentLegCd)
            Call TblLegLeaveCell(.Col, .Row)
        End If
        
        .SetFocus
    End With
End Sub

Private Sub cmdMove_Click()
    Call SpecMove
    tblLeg.SetFocus
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()

    Me.Show
    DoEvents
    
    With tblLeg
        BasicRowHeight = .RowHeight(1)
        BasicColWidth = .ColWidth(1)
        TableHeight = BasicRowHeight * .MaxRows
        TableWidth = BasicColWidth * .MaxCols
    End With

    Call QueryLegs
    If tabLeg.Tabs.Count > 0 Then
        CurrentLegCd = tabLeg.Tabs.Item(1).Key
        Call QueryLegInfo(CurrentLegCd)
        Call ReadSpecList(CurrentLegCd)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub tabLeg_Click()
    Dim LegCd As String
    
    tblLeg.ReDraw = False
        
    LegCd = tabLeg.SelectedItem.Key
'    If legCd = CurrentLegCd Then Exit Sub
    
    CurrentLegCd = LegCd
    
    Call QueryLegInfo(LegCd)
    Call ReadSpecList(LegCd)
    ClearSpcDetailInfo
    
    With tblLeg
        .Col = 1
        .Row = 1
        .Action = ActionActiveCell
        .SetFocus
    End With
    
    Call TblLegLeaveCell(1, 1)
    
    tblLeg.ReDraw = True
End Sub

Private Sub tblLeg_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewCol < 0 Or NewRow < 0 Then Exit Sub
    
    Call TblLegLeaveCell(NewCol, NewRow)
    lblRowCnt.Caption = NewRow
    lblColCnt.Caption = NewCol
End Sub

Private Sub QueryLegs()
    Dim i As Long
    Dim legCnt As Long
    Dim legList() As String
    Dim objSpcM As clsSpecManagement
    
    Set objSpcM = New clsSpecManagement
        
    legCnt = objSpcM.GetLegList(ObjSysInfo.BuildingCd, legList)
    
    Set objSpcM = Nothing
    
    cboLeg.Clear
    tabLeg.Tabs.Clear
    tblLeg.MaxCols = 0
    tblLeg.MaxRows = 0
    
    If legCnt < 1 Then Exit Sub
    
    For i = LBound(legList) To UBound(legList)
        tabLeg.Tabs.Add , legList(i), "Rack (" & legList(i) & ")"
        cboLeg.AddItem legList(i)
    Next i
    
End Sub

Private Sub QueryLegInfo(ByVal LegCd As String)
    Dim objSpcM As clsSpecManagement
    
    Set objSpcM = New clsSpecManagement
    If objSpcM.ReadLegInfo(ObjSysInfo.BuildingCd, LegCd) = True Then
        lblRowCnt.Caption = objSpcM.RowCnt
        lblColCnt.Caption = objSpcM.ColCnt
        lblRmk.Caption = objSpcM.rmk
    End If
    Set objSpcM = Nothing
    
    Call MakeLegShape(Val(lblRowCnt.Caption), Val(lblColCnt.Caption))
    
End Sub

Private Sub MakeLegShape(ByVal RowCnt As Long, ByVal ColCnt As Long)
    Dim i As Long
    Dim RowHeight As Double
    Dim ColWidth As Double
    
    
    If RowCnt = 0 Or ColCnt = 0 Then Exit Sub
    
    With tblLeg
        .MaxRows = RowCnt
        .MaxCols = ColCnt
        
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeTextWordWrap = True
        .TypeTextShadow = False
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        
        .BlockMode = False
        
        RowHeight = TableHeight / .MaxRows
        ColWidth = TableWidth / .MaxCols

        For i = 1 To .MaxRows
            .RowHeight(i) = RowHeight
        Next i
        
        For i = 1 To .MaxCols
            .ColWidth(i) = ColWidth
        Next i
        
    End With
End Sub

Private Function GetKeepHour() As Long
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSetDay(BC2_KEEP_HOUR)
    If DrRS Is Nothing Then
        GetKeepHour = 0
    End If
    With DrRS
        If .RecordCount < 1 Then
            GetKeepHour = 0
        Else
            GetKeepHour = .Fields("field1").value & ""
        End If
    End With
    Set DrRS = Nothing
    Set objcom003 = Nothing
End Function

Private Function Check72(ByVal spcyy As String, ByVal spcno As String, ByVal keephour As Long) As Boolean
    '72시간(보관시간)이   지났으면 true
    '                   안지났으면 false
    Dim DrRS As Recordset
    Dim objSpcM As clsSpecManagement
    Dim coldttm As Date
    Dim today As Date
    
    Set objSpcM = New clsSpecManagement
    Set DrRS = objSpcM.OpenBBS201(spcyy, spcno)
    Set objSpcM = Nothing
    
    If DrRS Is Nothing Then Exit Function
    
    today = GetSystemDate
    
    With DrRS
        If .RecordCount > 0 Then
            coldttm = Format(.Fields("coldt").value & "", "####-##-##") & " " & Format(.Fields("coltm").value & "", "00:00:00")
            If DateDiff("h", coldttm, today) > keephour Then
                Check72 = True
            End If
        End If
    End With
    Set DrRS = Nothing
End Function

Private Sub ReadSpecList(ByVal LegCd As String)
    Dim i As Long
    Dim spcno As String
    Dim stscd As String
    Dim keephour As Long
    
    Dim DrRS As Recordset
    Dim objSpcM As clsSpecManagement
        
    Dim objPrgBar As clsProgress
    
    Set objSpcM = New clsSpecManagement
    Set DrRS = objSpcM.GetSpcKeepSpace(ObjSysInfo.BuildingCd, LegCd)
    
    If DrRS Is Nothing Then Exit Sub
    If DrRS.RecordCount < 1 Then Exit Sub
    
    keephour = GetKeepHour
    
    Set objPrgBar = New clsProgress
'    Set objPrgBar.MyForm = Me
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = MainFrm.stsBar
    objPrgBar.Min = 1
    objPrgBar.Max = DrRS.RecordCount
    
    With tblLeg
        For i = 1 To DrRS.RecordCount
        
            objPrgBar.value = i
            
            stscd = DrRS.Fields("stscd").value & ""
            .Row = DrRS.Fields("rowno").value & ""
            .Col = DrRS.Fields("colno").value & ""
            
            If stscd = "1" Or stscd = "2" Then
                spcno = DrRS.Fields("spcyy").value & "" & "-" & DrRS.Fields("spcno").value & ""
'                .TypeTextShadow = True
                .BackColor = Me.BackColor
                If Check72(DrRS.Fields("spcyy").value & "", DrRS.Fields("spcno").value & "", keephour) = True Then
                    .ForeColor = RGB(255, 0, 0)
                Else
                    .ForeColor = RGB(0, 0, 0)
                End If
            Else
                .BackColor = RGB(255, 255, 255)
                .ForeColor = RGB(0, 0, 0)
                spcno = ""
'                .TypeTextShadow = False
            End If
            
            .value = spcno
            .FontStrikethru = (stscd = "2")

            DrRS.MoveNext
        Next i
    End With
    
    Set objPrgBar = Nothing
    
    Set DrRS = Nothing
End Sub

Private Sub TblLegLeaveCell(ByVal Col As Long, ByVal Row As Long)
    Dim value As String
    Dim coldt As String
    Dim spcyy As String
    Dim spcno As String
    
    Dim today As Date
    Dim colday As Date
    
    Dim Rs As Recordset
    Dim objSpcM As clsSpecManagement
    
    tblLeg.Row = Row
    tblLeg.Col = Col
    value = tblLeg.value
    
    If value = "" Then
        
        cmdMove.Enabled = False
        cmdExp.Enabled = False
        cmdExp2.Enabled = False
        
        ClearSpcDetailInfo
        Exit Sub
    End If
    
    value = Replace(value, vbNewLine, "-")
    
    coldt = medGetP(value, 1, "-")
    spcyy = medGetP(value, 1, "-")
    spcno = medGetP(value, 2, "-")
    
    
    Set objSpcM = New clsSpecManagement
    
    Set Rs = objSpcM.OpenBBS201(spcyy, spcno)
    If Rs Is Nothing Then
        cmdMove.Enabled = False
        cmdExp.Enabled = False
        cmdExp2.Enabled = False
        ClearSpcDetailInfo
        Exit Sub
    End If
    
    If Rs.RecordCount < 1 Then
    
        cmdMove.Enabled = False
        cmdExp.Enabled = False
        cmdExp2.Enabled = False
        
        Set Rs = Nothing
        ClearSpcDetailInfo
        Exit Sub
    End If
    
    
    cmdMove.Enabled = True
    cmdExp.Enabled = True
    cmdExp2.Enabled = True
    
    today = GetSystemDate
    colday = CDate(Format(Rs.Fields("coldt").value & "", "####-##-##") & " " & _
                   Format(Rs.Fields("coltm").value & "", "00:00:00"))
    
    lblSpcNo.Caption = spcyy & "-" & spcno
    lblColDtTm.Caption = Format(Rs.Fields("coldt").value & "", "####-##-##") & " " & _
                         Format(Mid(Rs.Fields("coltm").value & "", 1, 4), "00:00")
    lblColNm.Caption = GetEmpNm(Rs.Fields("colid").value & "")
    lblRcvDtTm.Caption = Format(Rs.Fields("rcvdt").value & "", "####-##-##") & " " & _
                         Format(Mid(Rs.Fields("rcvtm").value & "", 1, 4), "00:00")
    lblRcvNm.Caption = GetEmpNm(Rs.Fields("rcvid").value & "")
    lblPtNm.Caption = GetPtNm(Rs.Fields("ptid").value & "")
    lblPtId.Caption = Rs.Fields("ptid").value & ""
    lblLocalNm.Caption = GetLocalNm(Rs.Fields("localcd").value & "")
    lblPassTime.Caption = DateDiff("h", colday, today)
    
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    
    Set objSpcM = Nothing
End Sub

Private Sub ClearSpcDetailInfo()
    Call ICSPatientMark
    lblSpcNo.Caption = ""
    lblColDtTm.Caption = ""
    lblColNm.Caption = ""
    lblRcvDtTm.Caption = ""
    lblRcvNm.Caption = ""
    lblPtNm.Caption = ""
    lblPtId.Caption = ""
    lblLocalNm.Caption = ""
    lblPassTime.Caption = ""
End Sub

Private Sub SpecMove()
    '보관장소이동
    Dim spcyy As String
    Dim spcno As String
    Dim value As String
    
    Dim LegCd As String
    Dim RowNo As String
    Dim ColNo As String
    
    Dim objSpcM As clsSpecManagement
    
    LegCd = cboLeg.Text
    RowNo = cboRow.Text
    ColNo = cboCol.Text
    
    If LegCd = "" Then Exit Sub
    If RowNo = "" Then Exit Sub
    If ColNo = "" Then Exit Sub
    
    With tblLeg
        .Row = .ActiveRow
        .Col = .ActiveCol
        value = .value
        
'        value = Replace(value, vbNewLine, "-")
        spcyy = medGetP(value, 1, "-")
        spcno = medGetP(value, 2, "-")
    End With
    
    If spcyy = "" Then Exit Sub
    If spcno = "" Then Exit Sub
    
    Set objSpcM = New clsSpecManagement
    
    If objSpcM.SpecMove(ObjSysInfo.BuildingCd, spcyy, spcno, LegCd, RowNo, ColNo) = True Then
        Call QueryLegInfo(CurrentLegCd)
        Call ReadSpecList(CurrentLegCd)
        Call TblLegLeaveCell(tblLeg.ActiveCol, tblLeg.ActiveRow)
    End If
    
    Set objSpcM = Nothing
End Sub

Private Function Expire(ByVal spcyy As String, ByVal spcno As String, ByVal expfg As String) As Boolean
    Dim objSpcM As clsSpecManagement
    
    Set objSpcM = New clsSpecManagement
    Expire = objSpcM.Expire(ObjSysInfo.BuildingCd, spcyy, spcno, expfg)
    Set objSpcM = Nothing
End Function


