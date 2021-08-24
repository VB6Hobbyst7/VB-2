VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTLAMonitor 
   Appearance      =   0  '평면
   BackColor       =   &H00F8E4D8&
   Caption         =   "[TLA 모니터]"
   ClientHeight    =   12435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   12435
   ScaleWidth      =   22470
   StartUpPosition =   2  '화면 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame fraING 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5880
      TabIndex        =   8
      Top             =   180
      Width           =   11925
      Begin VB.CheckBox chkO 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Caption         =   "검사중"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   420
         TabIndex        =   16
         Top             =   570
         Value           =   1  '확인
         Width           =   2145
      End
      Begin VB.CheckBox chkOO 
         Appearance      =   0  '평면
         BackColor       =   &H0000FFFF&
         Caption         =   "재검"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2610
         TabIndex        =   15
         Top             =   570
         Value           =   1  '확인
         Width           =   2145
      End
      Begin VB.CheckBox chkR 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "결과"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   4800
         TabIndex        =   14
         Top             =   570
         Value           =   1  '확인
         Width           =   2145
      End
      Begin VB.CheckBox chkRR 
         Appearance      =   0  '평면
         BackColor       =   &H0000FF00&
         Caption         =   "재결"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   6990
         TabIndex        =   13
         Top             =   570
         Value           =   1  '확인
         Width           =   2145
      End
      Begin VB.CheckBox chkH 
         Appearance      =   0  '평면
         BackColor       =   &H00FF00FF&
         Caption         =   "에러"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   9180
         TabIndex        =   12
         Top             =   570
         Value           =   1  '확인
         Width           =   2145
      End
      Begin VB.Label lblO 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   420
         TabIndex        =   57
         Top             =   1350
         Width           =   2235
      End
      Begin VB.Label lblH 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   9240
         TabIndex        =   29
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Label lblRR 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   7050
         TabIndex        =   28
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Label lblR 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   4860
         TabIndex        =   27
         Top             =   1350
         Width           =   2145
      End
      Begin VB.Label lblOO 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   2640
         TabIndex        =   26
         Top             =   1350
         Width           =   2175
      End
      Begin VB.Label lblDate 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "TLA 모니터링중..."
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   6960
         TabIndex        =   22
         Top             =   210
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   120
         Picture         =   "Form1.frx":144A
         Top             =   210
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "검사진행현황"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   420
         TabIndex        =   17
         Top             =   180
         Width           =   2745
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         Height          =   1725
         Left            =   -630
         Top             =   540
         Width           =   12075
      End
   End
   Begin VB.Frame fraSet 
      Appearance      =   0  '평면
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  '없음
      Caption         =   " 조회기간설정 "
      ForeColor       =   &H80000008&
      Height          =   7245
      Left            =   9150
      TabIndex        =   23
      Top             =   2580
      Visible         =   0   'False
      Width           =   5595
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   62
         Top             =   5820
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648447
         Caption         =   "1 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtLimitSep 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1290
         TabIndex        =   53
         Text            =   "10"
         Top             =   3750
         Width           =   1605
      End
      Begin VB.TextBox txtLimit 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1290
         TabIndex        =   51
         Text            =   "50"
         Top             =   3270
         Width           =   1605
      End
      Begin VB.TextBox txtCOBTAT 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   45
         Text            =   "120"
         Top             =   2580
         Width           =   1605
      End
      Begin VB.TextBox txtAUTAT 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   43
         Text            =   "90"
         Top             =   2130
         Width           =   1605
      End
      Begin VB.TextBox txtARCTAT 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   40
         Text            =   "120"
         Top             =   1680
         Width           =   1605
      End
      Begin VB.TextBox txtSec 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1260
         TabIndex        =   35
         Text            =   "30"
         Top             =   150
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   435
         Left            =   3390
         TabIndex        =   31
         Top             =   900
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   69337089
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   435
         Left            =   1260
         TabIndex        =   32
         Top             =   900
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   69337089
         CurrentDate     =   40248
      End
      Begin HSCotrol.CButton cmdApply 
         Height          =   480
         Left            =   2670
         TabIndex        =   33
         Top             =   4590
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   847
         BackColor       =   16777152
         Caption         =   "Apply"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   480
         Left            =   3990
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4590
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   847
         BackColor       =   16777215
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   1
         Left            =   420
         TabIndex        =   63
         Top             =   6210
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648430
         Caption         =   "2 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   2
         Left            =   420
         TabIndex        =   64
         Top             =   6600
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648413
         Caption         =   "3 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   3
         Left            =   420
         TabIndex        =   65
         Top             =   6990
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648396
         Caption         =   "4 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   4
         Left            =   420
         TabIndex        =   66
         Top             =   7380
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648379
         Caption         =   "5 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   5
         Left            =   420
         TabIndex        =   67
         Top             =   7770
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12648362
         Caption         =   "6 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   6
         Left            =   2820
         TabIndex        =   68
         Top             =   5820
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640426
         Caption         =   "7 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   7
         Left            =   2820
         TabIndex        =   69
         Top             =   6210
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640443
         Caption         =   "8 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   8
         Left            =   2820
         TabIndex        =   70
         Top             =   6600
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640460
         Caption         =   "9 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   9
         Left            =   2820
         TabIndex        =   71
         Top             =   6990
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640477
         Caption         =   "10 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   10
         Left            =   2820
         TabIndex        =   72
         Top             =   7380
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640494
         Caption         =   "11 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   11
         Left            =   2820
         TabIndex        =   73
         Top             =   7770
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         CaptionBackColor=   12640511
         Caption         =   "12 Limit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSCotrol.CButton cmdColor 
         Height          =   480
         Left            =   150
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4590
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   847
         BackColor       =   16777215
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Limit 간격"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   210
         TabIndex        =   55
         Top             =   3885
         Width           =   945
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3030
         TabIndex        =   54
         Top             =   3810
         Width           =   555
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3030
         TabIndex        =   52
         Top             =   3330
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "대기 Limit"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   210
         TabIndex        =   50
         Top             =   3405
         Width           =   945
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   90
         X2              =   5415
         Y1              =   3120
         Y2              =   3135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   49
         Top             =   2670
         Width           =   555
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   48
         Top             =   2190
         Width           =   555
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "분"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   47
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "COB"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   1440
         TabIndex        =   46
         Top             =   2670
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "AU"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   1440
         TabIndex        =   44
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "ARC"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   42
         Top             =   1770
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "TAT Over"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   1815
         Width           =   930
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   90
         X2              =   5415
         Y1              =   1500
         Y2              =   1515
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3210
         TabIndex        =   39
         Top             =   1020
         Width           =   105
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   38
         Top             =   1020
         Width           =   780
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   90
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "조회주기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "초"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   36
         Top             =   180
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   90
         X2              =   5415
         Y1              =   690
         Y2              =   705
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   90
         Top             =   150
         Width           =   1140
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   120
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   90
         Top             =   3270
         Width           =   1140
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   90
         Top             =   3750
         Width           =   1140
      End
   End
   Begin VB.Frame fraTAT 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   15630
      TabIndex        =   10
      Top             =   180
      Width           =   6675
      Begin HSCotrol.CButton cmdExit 
         Height          =   570
         Left            =   5160
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   690
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1005
         BackColor       =   12632319
         Caption         =   "Exit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdRefresh 
         Height          =   570
         Left            =   2490
         TabIndex        =   20
         Top             =   690
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1005
         BackColor       =   16761087
         Caption         =   "Refresh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdSet 
         Height          =   570
         Left            =   3810
         TabIndex        =   30
         Top             =   690
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1005
         BackColor       =   16777152
         Caption         =   "설정"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "OVER TIME 갯수 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   36
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1380
         TabIndex        =   60
         Top             =   1500
         Width           =   4845
      End
      Begin VB.Label lblTAT 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "152"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   45.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   885
         Left            =   6570
         TabIndex        =   59
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         Height          =   885
         Left            =   6210
         Top             =   1350
         Width           =   1965
      End
      Begin VB.Label lblSec 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   24
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1200
         TabIndex        =   21
         Top             =   690
         Width           =   1245
      End
      Begin VB.Image Image3 
         Height          =   225
         Left            =   120
         Picture         =   "Form1.frx":1834
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "TAT Over 현황"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   420
         TabIndex        =   18
         Top             =   210
         Width           =   2745
      End
   End
   Begin VB.Frame fraSTB 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   7755
      Begin VB.Label lblStatus 
         BackStyle       =   0  '투명
         Caption         =   "TLA 모니터링중..."
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3720
         TabIndex        =   61
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label lblStandBy 
         BackStyle       =   0  '투명
         Caption         =   "대기 123 검체"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   56.25
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1005
         Left            =   30
         TabIndex        =   58
         Top             =   840
         Width           =   7305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "검사대기현황"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   150
         Width           =   2745
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   120
         Picture         =   "Form1.frx":1C1E
         Top             =   180
         Width           =   150
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         Height          =   1725
         Left            =   0
         Top             =   540
         Width           =   7455
      End
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   6975
      Left            =   1380
      TabIndex        =   3
      Top             =   2820
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CheckBox chkS 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1020
         TabIndex        =   25
         Top             =   1410
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Timer tmrMonitor 
         Left            =   2460
         Top             =   960
      End
      Begin VB.Timer tmrRefresh 
         Left            =   2940
         Top             =   990
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   390
         TabIndex        =   7
         Text            =   "30"
         Top             =   330
         Visible         =   0   'False
         Width           =   3885
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   750
         TabIndex        =   4
         Top             =   900
         Width           =   1425
      End
      Begin HSCotrol.CButton CButton1 
         Height          =   360
         Left            =   720
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2070
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   65280
         Caption         =   "검사중"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
      End
      Begin HSCotrol.CButton CButton2 
         Height          =   360
         Left            =   2010
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2070
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   16711935
         Caption         =   "검사완료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
      End
      Begin FPSpreadADO.fpSpread spdGrEq 
         Height          =   3105
         Left            =   480
         TabIndex        =   24
         Top             =   2910
         Width           =   2985
         _Version        =   524288
         _ExtentX        =   5265
         _ExtentY        =   5477
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   10
         SpreadDesigner  =   "Form1.frx":2008
      End
      Begin HSCotrol.CButton cmdDate 
         Height          =   360
         Left            =   3000
         TabIndex        =   56
         Top             =   1590
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         BackColor       =   16777152
         Caption         =   "조회기간설정"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
   End
   Begin FPSpreadADO.fpSpread spdSTB 
      Height          =   8895
      Left            =   180
      TabIndex        =   0
      Top             =   2550
      Width           =   5550
      _Version        =   524288
      _ExtentX        =   9790
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   4194368
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   5
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   16777215
      SpreadDesigner  =   "Form1.frx":243F
      UserResize      =   0
   End
   Begin FPSpreadADO.fpSpread spdING 
      Height          =   8895
      Left            =   5880
      TabIndex        =   1
      Top             =   2550
      Width           =   9540
      _Version        =   524288
      _ExtentX        =   16828
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   4
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "Form1.frx":2A90
      UserResize      =   0
   End
   Begin FPSpreadADO.fpSpread spdTAT 
      Height          =   8895
      Left            =   15630
      TabIndex        =   2
      Top             =   2550
      Width           =   6660
      _Version        =   524288
      _ExtentX        =   11748
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   5
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "Form1.frx":33A7
      UserResize      =   0
   End
End
Attribute VB_Name = "frmTLAMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngTime As Long

Private Sub cmdApply_Click()
    
    'tmrRefresh.Interval = txtSec.Text * 1000
    'tmrRefresh.Enabled = True
    
    lblDate.Caption = dtpStartDt.Value & "~" & dtpStopDt.Value
    
    Call WritePrivateProfileString("TLA", "REFRESH", txtSec.Text, App.Path & "\TLA.ini")
    Call WritePrivateProfileString("TLA", "ARCTAT", txtARCTAT.Text, App.Path & "\TLA.ini")
    Call WritePrivateProfileString("TLA", "AUTAT", txtAUTAT.Text, App.Path & "\TLA.ini")
    Call WritePrivateProfileString("TLA", "COBTAT", txtCOBTAT.Text, App.Path & "\TLA.ini")
    Call WritePrivateProfileString("TLA", "LIMIT", txtLimit.Text, App.Path & "\TLA.ini")
    Call WritePrivateProfileString("TLA", "LIMITSEP", txtLimitSep.Text, App.Path & "\TLA.ini")
    
    lblSec.Caption = txtSec.Text
    
End Sub

Private Sub cmdClose_Click()

    fraSet.Visible = False
    
    '-- 검사대기현황
    'Call GetData_STB
    
    '-- 검사진행현황
    'Call GetData_ING
    
    
End Sub

Private Sub cmdColor_Click()
    
    If fraSet.Height = 8385 Then
        fraSet.Height = 5385
    Else
        fraSet.Height = 8385
    End If
    
End Sub

'Private Sub cmdDate_Click()
'    If fraDateSet.Visible = True Then
'        fraDateSet.Visible = False
'    Else
'        fraDateSet.Visible = True
'        fraDateSet.ZOrder 0
'    End If
'
'End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdRefresh_Click()
    
    dtpStartDt.Value = Date
    dtpStopDt.Value = Date
    
    spdSTB.MaxRows = 0
    spdING.MaxRows = 0
    spdTAT.MaxRows = 0
    
    '-- 검사대기현황
    Call GetData_STB
    
    '-- 검사진행현황
    Call GetData_ING

    '-- TAT Over 현황
    Call GetTLAData_TAT

    
End Sub

Private Sub cmdSet_Click()
    
    If fraSet.Visible = True Then
        fraSet.Visible = False
    Else
        fraSet.Visible = True
        fraSet.ZOrder 0
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Width = 22800
    Me.Height = 13365
    
    '-- 초기화
    Call FrmInitial
    
    Call GetIni

    '-- DB 연결
    If Connect_PRServer Then
        lblStatus.Caption = "TLA 서버 연결 성공"
        tmrMonitor.Interval = 500
        tmrMonitor.Enabled = True
    Else
        lblStatus.Caption = "TLA 서버에 연결되지 않았습니다."
        Exit Sub
    End If
    
    
    '-- 그룹리스트조회
    Call GetGRPList
    
    '-- 장비리스트조회
    Call GetEQPList
    
    '-- 검사대기현황
    Call GetData_STB
    
    '-- 검사진행현황
    Call GetData_ING
    
    '-- TAT Over 현황
    Call GetTLAData_TAT
    
End Sub

'-- 장비리스트조회
Private Sub GetEQPList()
    Dim intRow      As Integer
    
    intRow = 0
    
          SQL = "Select DISTINCT GRPCD,EQPCD " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_EQP" & vbCr
    SQL = SQL & " Where LISCD < 1000" & vbCr
    SQL = SQL & " Order By GRPCD " & vbCr
 
    Call SetSQLData("02.장비리스트조회", SQL)
    
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        With spdGrEq
            .MaxRows = RS_Ser.RecordCount
            Do Until RS_Ser.EOF
                intRow = intRow + 1
                Call .SetText(1, intRow, Trim(RS_Ser.Fields("EQPCD")))
                Call .SetText(2, intRow, Trim(RS_Ser.Fields("GRPCD")))
                
                gEQPCD = gEQPCD & "'" & Trim(RS_Ser.Fields("EQPCD")) & "',"
                RS_Ser.MoveNext
            Loop
            .RowHeight(-1) = 20
            RS_Ser.Close
        End With
        
        If gEQPCD <> "" Then
            gEQPCD = Mid(gEQPCD, 1, Len(gEQPCD) - 1)
        End If
    End If
    
End Sub

'-- 그룹리스트조회
Private Sub GetGRPList()
    Dim intCol As Long
    
    intCol = 0
    
          SQL = "Select DISTINCT GRPCD " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_EQP" & vbCr
    SQL = SQL & " Where LISCD < 1000" & vbCr
    SQL = SQL & " Order By GRPCD " & vbCr
 
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    Call SetSQLData("01.그룹리스트조회", SQL)
    
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        With spdING
            
            .MaxCols = colHeader + RS_Ser.RecordCount
            
            Do Until RS_Ser.EOF
                intCol = intCol + 1
                
                .Col = colHeader + intCol
                .Row = -1
                .CellType = CellTypeStaticText
                .TypeEditCharSet = TypeEditCharSetASCII
                .TypeEditCharCase = TypeEditCharCaseSetNone
                
                .TypeHAlign = TypeHAlignCenter
                .TypeVAlign = TypeVAlignCenter
                .ColWidth(colHeader + intCol) = gWIDTH
                
                Call spdING.SetText(colHeader + intCol, 0, Trim(RS_Ser.Fields("GRPCD")))
                
                gGRPCD = gGRPCD & "'" & Trim(RS_Ser.Fields("GRPCD")) & "',"
                
                RS_Ser.MoveNext
            Loop
            
            RS_Ser.Close
        End With
        
        If gGRPCD <> "" Then
            gGRPCD = Mid(gGRPCD, 1, Len(gGRPCD) - 1)
        End If
    End If
    
End Sub

Private Sub GetData_STB()
    Dim intRow      As Long
    Dim strDtTm     As String
    Dim strDtTmRS   As String
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    
          SQL = "Select BCD,PNM,(SEX +'/'+ AGE)as SA,BLDTM, ODYYN as 당일, QUICKYN as 빠른" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
    SQL = SQL & " Where NOT STA IN ('E','O','R','H')" & vbCr
    SQL = SQL & "   And (ODYYN IS NOT NULL AND ODYYN <> '' OR QUICKYN IS NOT NULL AND QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
    SQL = SQL & " Order By BLDTM desc" & vbCr
 
    Call SetSQLData("03.검사대기조회", SQL)
    
    spdSTB.ReDraw = False
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            spdSTB.MaxRows = spdSTB.MaxRows + 1
            intRow = spdSTB.MaxRows
            
            strDtTmRS = Format(Trim(RS_Ser.Fields("BLDTM")), "MM/DD HH:MM")
            
            Call spdSTB.SetText(1, intRow, Format(Trim(RS_Ser.Fields("BLDTM")), "MM/DD HH:MM"))
            Call spdSTB.SetText(2, intRow, Trim(RS_Ser.Fields("BCD")))
            Call spdSTB.SetText(3, intRow, Trim(RS_Ser.Fields("PNM")) & "[" & Trim(RS_Ser.Fields("SA")) & "]")
            Call spdSTB.SetText(4, intRow, IIf(Trim(RS_Ser.Fields("당일")) <> "", "V", ""))
            Call spdSTB.SetText(5, intRow, IIf(Trim(RS_Ser.Fields("빠른")) <> "", "V", ""))
            
            If DateDiff("n", strDtTmRS, strDtTm) < gLimit Then
                '배경 - 하얀색
                spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = vbWhite
            Else
                If DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 1) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(0).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 2) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(1).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 3) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(2).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 4) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(3).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 5) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(4).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 6) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(5).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 7) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(6).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 8) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(7).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 9) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(8).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 10) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(9).CaptionBackColor
                ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 11) Then
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(10).CaptionBackColor
                Else
                    spdSTB.Row = intRow: spdSTB.Col = 1: spdSTB.BackColor = LmColor(11).CaptionBackColor
                End If
            End If
            
            RS_Ser.MoveNext
        Loop
        
        spdSTB.RowHeight(-1) = 30
    End If
    
    RS_Ser.Close
    
    lblStandBy.Caption = "대기 " & spdSTB.MaxRows & " 검체"
    'lblStandBy.Caption = spdSTB.MaxRows
    
    spdSTB.ReDraw = True
    
End Sub

Private Sub GetData_ING()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intCnt      As Integer
    Dim strGrpCd    As String
    Dim strCond     As String
    Dim strBarNo    As String
    Dim intO        As Integer
    Dim intOO       As Integer
    Dim intR        As Integer
    Dim intRR       As Integer
    Dim intH        As Integer
    Dim intE        As Integer
    
    Dim totO        As Integer
    Dim totOO       As Integer
    Dim totR        As Integer
    Dim totRR       As Integer
    Dim totH        As Integer
    Dim totE        As Integer
    
    Dim strDtTm     As String
    
    On Error Resume Next
    
    strCond = ""
    If chkO.Value = "1" Then
        strCond = strCond & "'O',"
    End If
    If chkOO.Value = "1" Then
        strCond = strCond & "'OO',"
    End If
    If chkR.Value = "1" Then
        strCond = strCond & "'R',"
    End If
    If chkRR.Value = "1" Then
        strCond = strCond & "'RR',"
    End If
    If chkH.Value = "1" Then
        strCond = strCond & "'H',"
    End If
    If chkS.Value = "1" Then
        strCond = strCond & "'S',"
    End If
    If strCond <> "" Then
        strCond = Mid(strCond, 1, Len(strCond) - 1)
    End If
        
''''          SQL = "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE)as SA,a.ADTM,b.EQPCD,b.STA" & vbCr  ',a.BLDTM
'''''    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_SENDLIST b " & vbCr
''''    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
''''    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
''''    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
''''    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
''''    If strCond <> "" Then
''''        SQL = SQL & "   And b.STA IN (" & strCond & ")" & vbCr
''''    End If
''''    SQL = SQL & " Order By a.ADTM DESC,a.BCD ASC, b.STA ASC,b.EQPCD ASC " & vbCr

'''          SQL = "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, (Select 'ERR' from TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
'''    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
'''    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
'''    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
'''    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
'''    If strCond <> "" Then
'''        SQL = SQL & "   And b.STA IN (" & strCond & ")" & vbCr
'''    End If
'''    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
'''    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
'''    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
'''    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
'''    SQL = SQL & "                    And BCD = b.BCD " & vbCr
'''    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
'''    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
'''    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
'''    SQL = SQL & " Order By a.ADTM DESC,a.BCD ASC, b.STA DESC,b.EQPCD ASC " & vbCr

          SQL = "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '1' as SEQ, (Select top 1 'ERR' from TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'O'" & vbCr
    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "                    And BCD = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
'    SQL = SQL & " and a.STA = 'H'"
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '2' as SEQ, (Select top 1 'ERR' from TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'R'" & vbCr
    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "                    And BCD = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    'SQL = SQL & " and a.STA = 'H'"
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '3' as SEQ, (Select top 1 'ERR' from TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'OO'" & vbCr
    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "                    And BCD = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    'SQL = SQL & " and a.STA = 'H'"
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '4' as SEQ, (Select top 1 'ERR' from TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'RR'" & vbCr
    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "                    And BCD = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    'SQL = SQL & " and a.STA = 'H'"
    SQL = SQL & " Order By a.ADTM DESC,a.BCD ASC,SEQ, b.STA DESC, b.EQPCD ASC " & vbCr
    

    Call SetSQLData("04.검사진행조회", SQL)

    spdING.ReDraw = False

    strGrpCd = ""
    strBarNo = ""
    totE = 0
    
    lblRR.Caption = 0
    lblOO.Caption = 0
    lblR.Caption = 0
    lblO.Caption = 0
    lblH.Caption = 0
'    lble.Caption = 0
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            'If Trim(RS_Ser.Fields("BCD")) = "170118509645" Then Stop
            'spdING.Action = ActionActiveCell
            If Trim(RS_Ser.Fields("ASTA")) <> "H" Then
                If strBarNo <> Trim(RS_Ser.Fields("BCD")) Then
                    If Trim(RS_Ser.Fields("ERR")) & "" = "ERR" Then
                        totE = totE + 1
                    End If
                    
                    '-- 카운트 올리기
                    If intRR > 0 Then
                        totRR = totRR + 1
                        lblRR.Caption = lblRR.Caption + 1
                    ElseIf intOO > 0 Then
                        totOO = totOO + 1
                        lblOO.Caption = lblOO.Caption + 1
                    ElseIf intR > 0 Then
                        totR = totR + 1
                        lblR.Caption = lblR.Caption + 1
                    ElseIf intO > 0 Then
                        totO = totO + 1
                        lblO.Caption = lblO.Caption + 1
                    End If
                    
                    
                    
                    intO = 0
                    intR = 0
                    intOO = 0
                    intRR = 0
                    'intE = 0
                    
                    spdING.MaxRows = spdING.MaxRows + 1
                    intRow = spdING.MaxRows
                    Call spdING.SetText(1, intRow, Trim(RS_Ser.Fields("EQPCD")))
                    Call spdING.SetText(2, intRow, Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM"))
                    
                    'strDtTm = Format(RS_Ser.Fields("ADTM"), "####-##-## ##:##:##")
                    'strDtTm = Format(strDtTm, "MM/DD HH:MM")
    
                    'Call spdING.SetText(2, intRow, strDtTm)
                    Call spdING.SetText(3, intRow, Trim(RS_Ser.Fields("BCD")))
                    Call spdING.SetText(4, intRow, Trim(RS_Ser.Fields("PNM")) & "[" & Trim(RS_Ser.Fields("SA")) & "]")
                    'spdING.Action = ActionActiveCell
                    'DoEvents
                    
                End If
                
                '-- 그룹명 찾기
                strGrpCd = ""
                For intCnt = 1 To spdGrEq.MaxRows
                    spdGrEq.Row = intCnt
                    spdGrEq.Col = 1
                    If Trim(spdGrEq.Text) = Trim(RS_Ser.Fields("EQPCD")) Then
                        spdGrEq.Col = 2
                        strGrpCd = Trim(spdGrEq.Text)
                        Exit For
                    End If
                Next
                
                
                For intCol = colHeader To spdING.MaxCols
                    spdING.Col = intCol
                    spdING.Row = 0
                    If Trim(spdING.Text) = strGrpCd Then
                        'Call spdING.SetText(intCol, intRow, Trim(RS_Ser.Fields("STA")))
                        'spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = vbGreen
                        If Trim(RS_Ser.Fields("ERR")) & "" = "ERR" Then
                            Call spdING.SetText(intCol, intRow, "에러")
                            spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HFF00FF
                        Else
                            Select Case Trim(RS_Ser.Fields("STA"))
                            Case "O":
                                        Call spdING.SetText(intCol, intRow, "검사중")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HC0FFFF
                                        intO = 1
                            Case "OO":
                                        Call spdING.SetText(intCol, intRow, "재검")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HFFFF&
                                        intOO = 1
                            Case "R":
                                        Call spdING.SetText(intCol, intRow, "결과")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HC0FFC0
                                        intR = 1
                            Case "RR":
                                        Call spdING.SetText(intCol, intRow, "재결")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HFF00&
                                        intRR = 1
                            Case "H":
                                        Call spdING.SetText(intCol, intRow, "완료")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HFF00FF
                                        intH = 1
                            Case "S":
                                        Call spdING.SetText(intCol, intRow, "S")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = vbRed
                            End Select
                        End If
                    End If
                Next
            End If
            
            strBarNo = Trim(RS_Ser.Fields("BCD"))
            RS_Ser.MoveNext
        Loop
        
        '-- 카운트 올리기
        If intRR > 0 Then
            totRR = totRR + 1
        ElseIf intOO > 0 Then
            totOO = totOO + 1
        ElseIf intR > 0 Then
            totR = totR + 1
        ElseIf intO > 0 Then
            totO = totO + 1
        End If
        
        spdING.RowHeight(-1) = 30
    End If
    
    RS_Ser.Close
    
    lblO.Caption = totO
    lblOO.Caption = totOO
    lblR.Caption = totR
    lblRR.Caption = totRR
    lblH.Caption = totH
    lblH.Caption = totE
    
    spdING.ReDraw = True
    
    DoEvents

End Sub

Private Sub GetTLAData_TAT()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim varTmp      As Variant
    Dim strDtTm     As String
    Dim strBldDtTm  As String
    Dim strBarNo    As String
    Dim strPatInfo  As String
    Dim strEqpCd    As String
    Dim strOverTm   As String
    Dim intCnt      As Integer
    
'    Dim intCnt      As Integer
'    Dim strGrpCd    As String
'    Dim strCond     As String
'    Dim intO        As Integer
'    Dim intOO       As Integer
'    Dim intR        As Integer
'    Dim intRR       As Integer
'    Dim intH        As Integer

    On Error Resume Next
    
    Call SetSQLData("05-1.TAT조회", "")

    spdTAT.ReDraw = False
    spdTAT.MaxRows = 0
    intCnt = 0
        
    For intRow = 1 To spdING.MaxRows
        Call spdING.GetText(2, intRow, varTmp): strBldDtTm = varTmp
        Call spdING.GetText(3, intRow, varTmp): strBarNo = varTmp
        Call spdING.GetText(4, intRow, varTmp): strPatInfo = varTmp
        
        For intCol = 5 To 7
            Call spdING.GetText(intCol, intRow, varTmp): strEqpCd = varTmp
            If strEqpCd <> "" Then
                Call spdING.GetText(intCol, 0, varTmp): strEqpCd = varTmp
                
                      SQL = "Select INDTM " & vbCr
                SQL = SQL & "  From " & gDB & "..TB_EVENT " & vbCr
                SQL = SQL & " Where BCD = '" & strBarNo & "'" & vbCr
                SQL = SQL & "   And GRPCD = '" & strEqpCd & "'" & vbCr
                'SQL = SQL & "   And EQPCD LIKE '%" & strEqpCd & "%'" & vbCr
                'SQL = SQL & "   And STA <> 'H' " & vbCr
                SQL = SQL & "   And STA IN ('O','R') " & vbCr
                SQL = SQL & "   And SEQ = (Select MAX(SEQ) FROM " & gDB & "..TB_EVENT "
                SQL = SQL & "               Where BCD = '" & strBarNo & "'" & vbCr
                'SQL = SQL & "                 And EQPCD LIKE '%" & strEqpCd & "%'" & vbCr
                SQL = SQL & "                 And GRPCD = '" & strEqpCd & "'" & vbCr
                'SQL = SQL & "                 And STA <> 'H' " & vbCr
                SQL = SQL & "                 And STA IN ('O','R') " & vbCr
                SQL = SQL & "               Group By BCD) " & vbCr
                
                Call SetSQLData("05.TAT조회", SQL)
                
                Set RS_Ser = Cn_Ser.Execute(SQL)
                If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
                    Do Until RS_Ser.EOF
                        'strDtTm = CStr(Format(RS_Ser.Fields("INDTM") & "", "yyyy-mm-dd hh:mm:ss"))
                        strDtTm = Format(RS_Ser.Fields("INDTM"), "####-##-## ##:##:##")
                        strDtTm = Format(strDtTm, "MM/DD HH:MM")
                        strOverTm = DateDiff("n", strBldDtTm, strDtTm)
                        If intCol = 5 Then  'ARC
                            If strOverTm > gTatARC Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo)
                                Call spdTAT.SetText(3, intCnt, strPatInfo)
                                Call spdTAT.SetText(4, intCnt, "ARC")
                                'Call spdTAT.SetText(5, intCnt, strOverTm)
                                Call spdTAT.SetText(5, intCnt, strOverTm & " (+" & strOverTm - gTatARC & "m)")
                            End If
                        ElseIf intCol = 6 Then  'AU
                            If strOverTm > gTatAU Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo)
                                Call spdTAT.SetText(3, intCnt, strPatInfo)
                                Call spdTAT.SetText(4, intCnt, "AU")
                                'Call spdTAT.SetText(5, intCnt, strOverTm)
                                Call spdTAT.SetText(5, intCnt, strOverTm & " (+" & strOverTm - gTatAU & "m)")
                            End If
                        ElseIf intCol = 7 Then  'COB
                            If strOverTm > gTatCOB Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo)
                                Call spdTAT.SetText(3, intCnt, strPatInfo)
                                Call spdTAT.SetText(4, intCnt, "COB")
                                'Call spdTAT.SetText(5, intCnt, strOverTm)
                                Call spdTAT.SetText(5, intCnt, strOverTm & " (+" & strOverTm - gTatCOB & "m)")
                            End If
                        End If
                        RS_Ser.MoveNext
                    Loop
                End If
                RS_Ser.Close
            End If
        Next
    Next

    spdTAT.RowHeight(-1) = 30
    lblTAT.Caption = intCnt
    
    
    spdTAT.ReDraw = True

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    
    spdSTB.Width = Me.ScaleWidth * 0.25
    spdSTB.Height = Me.ScaleHeight - 3000
    fraSTB.Width = Me.ScaleWidth * 0.25
    
    spdING.Left = spdSTB.Left + spdSTB.Width + 100
    spdING.Width = Me.ScaleWidth * 0.4
    spdING.Height = Me.ScaleHeight - 3000
    fraING.Left = spdSTB.Left + spdSTB.Width + 100
    fraING.Width = Me.ScaleWidth * 0.4
    
    spdTAT.Left = spdING.Left + spdING.Width + 100
    spdTAT.Width = Me.ScaleWidth * 0.32
    spdTAT.Height = Me.ScaleHeight - 3000
    
    fraTAT.Left = spdING.Left + spdING.Width + 100
    fraTAT.Width = Me.ScaleWidth * 0.32
    
    DoEvents
    
End Sub



'Private Sub lblSec_Click()
'
'    If fraSecSet.Visible = False Then
'        fraSecSet.Visible = True
'    Else
'        fraSecSet.Visible = False
'    End If
'
'End Sub


Private Sub tmrMonitor_Timer()
    
    lblStatus.Caption = "TLA 모니터링중..."
    If lblStatus.Visible = True Then
        lblStatus.Visible = False
    Else
        lblStatus.Visible = True
    End If
    
End Sub

Private Sub FrmInitial()
    
    spdSTB.MaxRows = 0
    spdING.MaxRows = 0
    spdTAT.MaxRows = 0
    spdGrEq.MaxRows = 0
    spdGrEq.Visible = False
    
    dtpStartDt.Value = Date
    dtpStopDt.Value = Date
        
    lblDate.Caption = dtpStartDt.Value & "~" & dtpStopDt.Value
    lblStandBy.Caption = ""
    lblTAT.Caption = ""
    
    txtBarcode.Text = ""
    
    lblO.Caption = ""
    lblOO.Caption = ""
    lblR.Caption = ""
    lblRR.Caption = ""
    lblH.Caption = ""
    
    fraSet.Height = 5385
    
End Sub

Private Sub GetIni()
    Dim DB_Tmp As String * 100

    DB_Tmp = ""
    
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "IP", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gIP = Trim(txtTemp.Text)
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "DB", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gDB = Trim(txtTemp.Text)
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "UID", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gUID = Trim(txtTemp.Text)
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "PWD", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gPWD = Trim(txtTemp.Text)

    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "WIDTH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gWIDTH = Trim(txtTemp.Text)

    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "REFRESH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    lblSec.Caption = Trim(txtTemp.Text)
    txtSec.Text = Trim(txtTemp.Text)
    lngTime = Trim(txtTemp.Text)
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "ARCTAT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtARCTAT.Text = Trim(txtTemp.Text)
    gTatARC = txtARCTAT.Text
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "AUTAT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtAUTAT.Text = Trim(txtTemp.Text)
    gTatAU = txtAUTAT.Text
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "COBTAT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtCOBTAT.Text = Trim(txtTemp.Text)
    gTatCOB = txtCOBTAT.Text
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "LIMIT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtLimit.Text = Trim(txtTemp.Text)
    gLimit = txtLimit.Text
    
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "LIMITSEP", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtLimitSep.Text = Trim(txtTemp.Text)
    gLimitS = txtLimitSep.Text
    
    
'    tmrRefresh.Interval = 60000
    tmrRefresh.Interval = lblSec * 1000
    tmrRefresh.Enabled = True

    'Call WritePrivateProfileString("config", "gPort", gSetup.gPort, App.Path & "\Interface.ini")
    
    'txtBarcode.SetFocus

End Sub

'-- 주운영 서버
Public Function Connect_PRServer() As Boolean

    Connect_PRServer = False
        
On Error GoTo errFind
    
    Set Cn_Ser = New ADODB.Connection
    
    With Cn_Ser
        .ConnectionTimeout = 25
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = gIP      '"172.17.98.101"
        .Properties("Initial Catalog").Value = gDB  '"SNUBH_TLA"
        .Properties("User ID").Value = gUID         '"sa"
        .Properties("Password").Value = gPWD        '"sansoft@04"
        .Open
    End With
    
    Connect_PRServer = True
    
    Exit Function
 
errFind:
    
    Connect_PRServer = False
    
End Function

Private Sub tmrRefresh_Timer()
    
    Call cmdRefresh_Click
    
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    If KeyAscii = vbKeyReturn Then
        With spdING
            For intRow = 1 To .DataRowCnt
                .Row = intRow
                .Col = 3
                If Trim(.Text) = txtBarcode.Text Then
                    .Row = intRow
                    .Col = 3
                    .Action = ActionActiveCell
                    Exit For
                End If
            Next
        End With
    End If
    
End Sub
