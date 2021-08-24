VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS204 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검체추가요청서 작성"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "frmBBS204.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8010
   StartUpPosition =   1  '소유자 가운데
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   75
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3495
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  처방 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1500
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  환자 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   45
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  접수 번호"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   23
      Top             =   285
      Width           =   7875
      Begin VB.ComboBox cboRsn 
         Height          =   300
         ItemData        =   "frmBBS204.frx":076A
         Left            =   5820
         List            =   "frmBBS204.frx":0774
         Style           =   2  '드롭다운 목록
         TabIndex        =   26
         Top             =   450
         Width           =   1755
      End
      Begin VB.TextBox txtAccNo 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1260
         TabIndex        =   25
         Top             =   420
         Width           =   1410
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   195
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   420
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   4680
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   420
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
         Caption         =   "추가사유"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      Height          =   1755
      Left            =   75
      TabIndex        =   13
      Top             =   3735
      Width           =   7875
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   90
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   750
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
         Caption         =   "수량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   12
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   360
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
         Caption         =   "처방코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   13
         Left            =   2715
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   750
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
         Caption         =   "처방일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   14
         Left            =   2715
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
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
         Caption         =   "처방명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   15
         Left            =   5295
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   750
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
         Caption         =   "예정일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   16
         Left            =   90
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "수혈사유"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   17
         Left            =   2715
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "상태"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   18
         Left            =   5295
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "DC여부"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdCd 
         Height          =   360
         Left            =   1170
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdNm 
         Height          =   360
         Left            =   3780
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   3960
         _ExtentX        =   6985
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblUnitQty 
         Height          =   360
         Left            =   1170
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   360
         Left            =   3780
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReqDt 
         Height          =   360
         Left            =   6360
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   750
         Width           =   1380
         _ExtentX        =   2434
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTransReason 
         Height          =   360
         Left            =   1170
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatus 
         Height          =   360
         Left            =   3780
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDC 
         Height          =   360
         Left            =   6360
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         BackColor       =   14411494
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      Height          =   1755
      Left            =   75
      TabIndex        =   3
      Top             =   1740
      Width           =   7875
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   90
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "병 동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   2715
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "업무구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   5295
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1140
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
         Caption         =   "입원일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   90
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   750
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   90
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
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
         Index           =   6
         Left            =   2715
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   750
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
         Index           =   1
         Left            =   5295
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   360
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
         Caption         =   "성/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   2715
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   360
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
         Caption         =   "성 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   5295
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   750
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
      Begin MedControls1.LisLabel lblPtId 
         Height          =   360
         Left            =   1185
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   3780
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   360
         Left            =   6360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   360
         Left            =   1185
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoct 
         Height          =   360
         Left            =   3780
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   360
         Left            =   6360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   360
         Left            =   1185
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBuss 
         Height          =   360
         Left            =   4110
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1065
         _ExtentX        =   1879
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBedInDt 
         Height          =   360
         Left            =   6360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1395
         _ExtentX        =   2461
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBussCd 
         Height          =   360
         Left            =   3780
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1140
         Width           =   315
         _ExtentX        =   556
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
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   6615
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   5625
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   5295
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   5625
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   3975
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   5625
      Width           =   1320
   End
End
Attribute VB_Name = "frmBBS204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngAccNo As Long


Private Sub cmdClear_Click()
    Clear
    txtAccNo = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    Call BagGroundDisplay
    cboRsn.SetFocus
End Sub

Private Sub Form_Load()
    Call Get_Accdt
    Call Clear
    cmdSave.Enabled = True
End Sub
Private Sub Get_Accdt()
    Dim objNumbers As New clsBBSNumbers

    With objNumbers                 '접수 일자의 형식을 가져온다.(AccDt에 저장할 날짜)
        lngAccNo = Len(.Get_AccdtFormat)
    End With
    Set objNumbers = Nothing
End Sub
Private Sub Clear()
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblDoct.Caption = ""
    lblSexAge.Caption = ""
    lblDOB.Caption = ""
    lblWardNm.Caption = ""
    lblDeptNm.Caption = ""
    lblBussCd.Caption = ""
    lblBuss.Caption = ""
    lblBedInDt.Caption = ""
    lblDC.Caption = ""
    lblOrdCd.Caption = ""
    lblOrdNm.Caption = ""
    lblUnitQty.Caption = ""
    lblOrdDt.Caption = ""
    lblReqDt.Caption = ""
    lblTransReason.Caption = ""
    lblStatus.Caption = ""
    cboRsn.ListIndex = -1
    Call ICSPatientMark
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Function AccNoLen_Chk(ByVal accno As String) As Boolean
    
    If Len(accno) <= lngAccNo Then
        MsgBox "접수번호 형식과 일치 하지 않습니다. " & vbNewLine _
               & "접수번호는" & lngAccNo & " 자리 이상입니다.", vbCritical + vbOKOnly, Me.Caption
        AccNoLen_Chk = False
    Else
        AccNoLen_Chk = True
    End If
    
End Function

Private Sub BagGroundDisplay()
    '화면에 조회정보  Display
    Dim objGetSql       As clsGetSqlStatement
    Dim objTransReason  As clsQueryOrder
    Dim Rs              As New Recordset
    Dim strReason       As String
    Dim bedindt         As String
    Dim strTmp          As String
    Dim i As Integer
    
    
    If AccNoLen_Chk(txtAccNo) = False Then Exit Sub
        
    Set objGetSql = New clsGetSqlStatement
    
    Set Rs = objGetSql.Get_BagGround(txtAccNo, lngAccNo)
    
    
    With Rs
        If .RecordCount < 1 Then
            MsgBox "해당조건의 자료가 없습니다.확인후 입력하세요.", vbInformation + vbOKOnly, Me.Caption
            Clear
            cmdSave.Enabled = False
        Else
            
            bedindt = IIf(Len(.Fields("bedindt").value & "") = 8, .Fields("bedindt").value & "", "")
            
            lblPtNm.Caption = .Fields("ptnm").value & ""
            lblDeptNm.Caption = .Fields("deptcd").value & ""
            lblWardNm.Caption = .Fields("wardid").value & "" & "-" & _
                                .Fields("hosilid").value & "" & "-" & _
                                .Fields("bedid").value & ""
            lblDoct.Caption = GetDoctNm(.Fields("orddoct").value & "")
            lblPtId.Caption = .Fields("ptid").value & ""
            lblBussCd.Caption = .Fields("bussdiv").value & ""
            lblBedInDt.Caption = IIf(bedindt <> "", Format(bedindt, "####-##-##"), "")
            
            strTmp = SDA_String(.Fields("ssn").value & "")
            lblDOB.Caption = medGetP(strTmp, 2, COL_DIV)
            lblSexAge.Caption = medGetP(strTmp, 1, COL_DIV) & "/" & medGetP(strTmp, 3, COL_DIV)
            
            lblOrdDt.Caption = Format(.Fields("orddt").value & "", "####-##-##")
            lblOrdCd.Caption = .Fields("ordcd").value & ""
            lblOrdNm.Caption = .Fields("testnm").value & ""
            lblReqDt.Caption = Format(.Fields("reqdt").value & "", "####-##-##")
            lblUnitQty.Caption = .Fields("unitqty").value & ""
            lblDC.Caption = IIf(.Fields("dcfg").value & "" = "1", "Y", "")
            
            Select Case .Fields("bussdiv").value & ""
                Case "1": lblBuss.Caption = "외래"
                Case "2": lblBuss.Caption = "입원"
                Case "3": lblBuss.Caption = "응급"
            End Select
            
            Set objTransReason = New clsQueryOrder
            
            strReason = objTransReason.GetTransReason(lblPtId.Caption, Trim(Rs.Fields("orddt").value & ""), Trim(Rs.Fields("ordno").value & ""))
            lblTransReason.Caption = strReason
            
            If TRANS_REQUIRE_USED Then
                Select Case .Fields("stscd").value & ""
                    Case BBSOrdStatus.stsORDER:     lblStatus.Caption = "처방"
                    Case BBSOrdStatus.stsCOLLECT:   lblStatus.Caption = "채혈"
                    Case BBSOrdStatus.stsACCESS:    lblStatus.Caption = "접수"
                    Case BBSOrdStatus.stsINPROCESS: lblStatus.Caption = "검사중"
                    Case Else:                      lblStatus.Caption = ""
                End Select
            Else
                Select Case .Fields("stscd").value & ""
                    Case BBSOrderStatus.stsORDER:     lblStatus.Caption = "처방"
                    Case BBSOrderStatus.stsCOLLECT:   lblStatus.Caption = "채혈"
                    Case BBSOrderStatus.stsACCESS:    lblStatus.Caption = "접수"
                    Case BBSOrderStatus.stsINPROCESS: lblStatus.Caption = "검사중"
                    Case Else:                        lblStatus.Caption = ""
                    
                End Select
            End If
            cmdSave.Enabled = True
            Set objTransReason = Nothing
        End If
    End With
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    Set Rs = Nothing
    Set objGetSql = Nothing

End Sub
Private Sub cmdSave_Click()
    Dim PtId As String
    Dim reqdt As String
    Dim reqtm As String
    Dim reqid As String
    Dim accdt As String
    Dim accno As String
    Dim rsncd As String
    Dim busidiv As String
    Dim bedindt As String
    Dim orddt As String
    Dim DeptCd As String
    Dim wardid As String
    Dim donefg As String
    Dim spcyy As String
    Dim spcno As String
    
    If cboRsn.ListIndex < 0 Then
        MsgBox "추가요청 사유를 선택하세요", vbCritical + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Dim objbegin As New clsBeginTrans
    
    PtId = lblPtId.Caption
    reqid = CStr(ObjSysInfo.EmpId)
    reqdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    reqtm = Format(GetSystemDate, "hhmmss")
    accdt = medGetP(txtAccNo, 1, "-")
    accno = medGetP(txtAccNo, 2, "-")
    rsncd = cboRsn.ListIndex
    busidiv = lblBussCd.Caption
    bedindt = IIf(lblBedInDt.Caption <> "", Replace(lblBedInDt.Caption, "-", ""), "")
    orddt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    DeptCd = lblDeptNm.Caption
    wardid = medGetP(lblWardNm.Caption, 1, "-")
    
    If TRANS_REQUIRE_USED Then
        donefg = BBSOrdStatus.stsORDER
    Else
        donefg = BBSOrderStatus.stsORDER
    End If
    
    With objbegin
        If .Set_InsertSpcAdd(PtId, reqdt, reqtm, accdt, accno, rsncd, busidiv, bedindt, orddt, DeptCd, wardid, donefg, reqid) = False Then
            MsgBox "검체추가 요청오류입니다.", vbCritical + vbOKOnly, "검체추가요청오류"
        End If
    End With
    
    Call Clear
    txtAccNo = ""
    cmdSave.Enabled = False
    Set objbegin = Nothing
End Sub


Private Sub txtAccNo_GotFocus()
    txtAccNo.tag = txtAccNo
    txtAccNo.SelStart = 0
    txtAccNo.SelLength = Len(txtAccNo)
End Sub

Private Sub txtAccNo_Change()
    Dim lngLen As Long
    
    With txtAccNo
        lngLen = Len(Trim(.Text))
        If lngLen = lngAccNo Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)

    If Len(txtAccNo) <> lngAccNo Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtAccNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If

End Sub

Private Sub txtAccNo_LostFocus()
    If txtAccNo = "" Then
        Clear
    Else
        If txtAccNo.tag = txtAccNo Then
            Exit Sub
        Else
            BagGroundDisplay
        End If
    End If
End Sub





