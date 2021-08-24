VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
      Begin VB.Label lblTotCnt 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "당일 : 0 개 빠른 : 12개"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3210
         TabIndex        =   92
         Top             =   210
         Width           =   6975
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
         TabIndex        =   56
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   1350
         Width           =   2175
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
      Height          =   8505
      Left            =   9150
      TabIndex        =   22
      Top             =   2580
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtTATTAT 
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
         TabIndex        =   88
         Text            =   "120"
         Top             =   2850
         Width           =   1605
      End
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   0
         Left            =   2070
         TabIndex        =   74
         Top             =   5820
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1770
         Top             =   4890
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   61
         Top             =   5820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16777215
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
         TabIndex        =   52
         Text            =   "10"
         Top             =   3990
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
         TabIndex        =   50
         Text            =   "50"
         Top             =   3510
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
         TabIndex        =   44
         Text            =   "120"
         Top             =   2460
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
         TabIndex        =   42
         Text            =   "90"
         Top             =   2070
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
         TabIndex        =   39
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
         TabIndex        =   34
         Text            =   "30"
         Top             =   150
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   435
         Left            =   3390
         TabIndex        =   30
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
         Format          =   21364737
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   435
         Left            =   1260
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
         Format          =   21364737
         CurrentDate     =   40248
      End
      Begin HSCotrol.CButton cmdApply 
         Height          =   480
         Left            =   2670
         TabIndex        =   32
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
         TabIndex        =   33
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
         TabIndex        =   62
         Top             =   6210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16777215
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
         TabIndex        =   63
         Top             =   6600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16777215
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
         TabIndex        =   64
         Top             =   6990
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   65535
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
         TabIndex        =   65
         Top             =   7380
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   65535
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
         TabIndex        =   66
         Top             =   7770
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   65535
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
         TabIndex        =   67
         Top             =   5820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   8438015
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
         TabIndex        =   68
         Top             =   6210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   8438015
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
         TabIndex        =   69
         Top             =   6600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16761087
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
         TabIndex        =   70
         Top             =   6990
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16761087
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
         TabIndex        =   71
         Top             =   7380
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   16711935
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
         TabIndex        =   72
         Top             =   7770
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         CaptionBackColor=   255
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
         Left            =   420
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   4590
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   847
         BackColor       =   16777215
         Caption         =   "Color Set"
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   1
         Left            =   2070
         TabIndex        =   75
         Top             =   6210
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   2
         Left            =   2070
         TabIndex        =   76
         Top             =   6600
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   3
         Left            =   2070
         TabIndex        =   77
         Top             =   6990
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   4
         Left            =   2070
         TabIndex        =   78
         Top             =   7380
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   5
         Left            =   2070
         TabIndex        =   79
         Top             =   7770
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   6
         Left            =   4500
         TabIndex        =   80
         Top             =   5820
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   7
         Left            =   4500
         TabIndex        =   81
         Top             =   6210
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   8
         Left            =   4500
         TabIndex        =   82
         Top             =   6600
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   9
         Left            =   4500
         TabIndex        =   83
         Top             =   6990
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   10
         Left            =   4500
         TabIndex        =   84
         Top             =   7380
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   11
         Left            =   4500
         TabIndex        =   85
         Top             =   7770
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   688
         BackColor       =   16777215
         Caption         =   "Set"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "TAT"
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
         Index           =   9
         Left            =   1440
         TabIndex        =   90
         Top             =   2970
         Width           =   465
      End
      Begin VB.Label Label10 
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
         TabIndex        =   89
         Top             =   2940
         Width           =   555
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
         TabIndex        =   54
         Top             =   4125
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
         TabIndex        =   53
         Top             =   4050
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
         TabIndex        =   51
         Top             =   3570
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
         TabIndex        =   49
         Top             =   3645
         Width           =   945
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   90
         X2              =   5415
         Y1              =   3360
         Y2              =   3375
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
         TabIndex        =   48
         Top             =   2550
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
         TabIndex        =   47
         Top             =   2130
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   2580
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
         TabIndex        =   43
         Top             =   2190
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         Top             =   3510
         Width           =   1140
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00CEBE73&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   435
         Left            =   90
         Top             =   3990
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
         Left            =   6450
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
         TabIndex        =   29
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
      Begin HSCotrol.CButton cmdExcel 
         Height          =   570
         Left            =   5130
         TabIndex        =   87
         Top             =   690
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1005
         BackColor       =   12648384
         Caption         =   "엑셀저장"
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
         Left            =   3900
         TabIndex        =   91
         Top             =   300
         Width           =   2475
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
         TabIndex        =   59
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
         TabIndex        =   58
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
      Begin VB.Label lblCnt 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "당일 : 0 개 빠른 : 12개"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   86
         Top             =   2010
         Width           =   4725
      End
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
         TabIndex        =   60
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
         TabIndex        =   57
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
      Left            =   540
      TabIndex        =   3
      Top             =   4230
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CheckBox chkS 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1020
         TabIndex        =   24
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
         Height          =   1695
         Left            =   480
         TabIndex        =   23
         Top             =   2910
         Width           =   2715
         _Version        =   524288
         _ExtentX        =   4789
         _ExtentY        =   2990
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
         TabIndex        =   55
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
      Begin FPSpreadADO.fpSpread spdTOT 
         Height          =   1095
         Left            =   120
         TabIndex        =   93
         Top             =   5040
         Width           =   4950
         _Version        =   524288
         _ExtentX        =   8731
         _ExtentY        =   1931
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
         MaxCols         =   7
         MaxRows         =   499
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   16777215
         SpreadDesigner  =   "Form1.frx":243F
         UserResize      =   0
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
      MaxCols         =   7
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   16777215
      SpreadDesigner  =   "Form1.frx":2B34
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
      MaxCols         =   6
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "Form1.frx":3229
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
      MaxCols         =   7
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "Form1.frx":3BE4
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

Private Sub CButton3_Click()

End Sub

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

Private Sub cmdColorSet_Click(Index As Integer)
    Dim GetColor
    
    CommonDialog1.ShowColor
    
    GetColor = CommonDialog1.Color
    LmColor(Index).CaptionBackColor = GetColor
    
        
    Select Case Index
        Case 0:     Call WritePrivateProfileString("COLOR", "LV1", CStr(GetColor), App.Path & "\TLA.ini")
        Case 1:     Call WritePrivateProfileString("COLOR", "LV2", CStr(GetColor), App.Path & "\TLA.ini")
        Case 2:     Call WritePrivateProfileString("COLOR", "LV3", CStr(GetColor), App.Path & "\TLA.ini")
        Case 3:     Call WritePrivateProfileString("COLOR", "LV4", CStr(GetColor), App.Path & "\TLA.ini")
        Case 4:     Call WritePrivateProfileString("COLOR", "LV5", CStr(GetColor), App.Path & "\TLA.ini")
        Case 5:     Call WritePrivateProfileString("COLOR", "LV6", CStr(GetColor), App.Path & "\TLA.ini")
        Case 6:     Call WritePrivateProfileString("COLOR", "LV7", CStr(GetColor), App.Path & "\TLA.ini")
        Case 7:     Call WritePrivateProfileString("COLOR", "LV8", CStr(GetColor), App.Path & "\TLA.ini")
        Case 8:     Call WritePrivateProfileString("COLOR", "LV9", CStr(GetColor), App.Path & "\TLA.ini")
        Case 9:     Call WritePrivateProfileString("COLOR", "LV10", CStr(GetColor), App.Path & "\TLA.ini")
        Case 10:    Call WritePrivateProfileString("COLOR", "LV11", CStr(GetColor), App.Path & "\TLA.ini")
        Case 11:    Call WritePrivateProfileString("COLOR", "LV12", CStr(GetColor), App.Path & "\TLA.ini")
    End Select
    
End Sub

Private Sub cmdExcel_Click()
    Dim sFileName As String
    
    Call GetData_Tot_Print

'    If spdING.DataRowCnt < 1 Then
'        MsgBox "저장할 자료가 없습니다.", , "알 림"
'        Exit Sub
'    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        
        If SaveExcel(sFileName) = True Then
            MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
        Else
            MsgBox "엑셀 저장실패", vbOKOnly + vbInformation, Me.Caption
        End If
'    End If
    
End Sub

Private Function SaveExcel(Filename As String) As Boolean

On Error GoTo Rst
    
    SaveExcel = False
    
    ' Excel Object Library 와 연결합니다.
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)

'    For iRow = 0 To spdSTB.DataRowCnt
'        For iCol = 1 To spdSTB.DataColCnt
'            spdSTB.Row = iRow
'            spdSTB.Col = iCol
'            xlSheet.Cells(iRow + 1, iCol) = CStr(spdSTB.Text)
''            xlSheet.Type = 1
'        Next iCol
'    Next iRow

    Clipboard.Clear
    
    spdSTB.Col = 1:
    spdSTB.Col2 = spdSTB.MaxCols
    spdSTB.Row = 0
    spdSTB.Row2 = spdSTB.MaxRows
    
    Clipboard.SetText spdSTB.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = "검사대기현황"
    
    Set xlSheet = xlBook.Worksheets(2)
'    xlBook.Worksheets("Sheet2").Name = "검사진행현황"
'    Set xlSheet = xlBook.Worksheets("검사진행현황")

'    For iRow = 0 To spdING.DataRowCnt
'        For iCol = 1 To spdING.DataColCnt
'            spdING.Row = iRow
'            spdING.Col = iCol
'            xlSheet.Cells(iRow + 1, iCol) = CStr(spdING.Text)
'        Next iCol
'    Next iRow

    Clipboard.Clear
    
    spdING.Col = 2:
    spdING.Col2 = spdING.MaxCols
    spdING.Row = 0
    spdING.Row2 = spdING.MaxRows
    
    Clipboard.SetText spdING.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = "검사진행현황"

    Set xlSheet = xlBook.Worksheets(3)
'    xlBook.Worksheets("Sheet3").Name = "TATOver현황"
'    Set xlSheet = xlBook.Worksheets("TATOver현황")

'    For iRow = 0 To spdTAT.DataRowCnt
'        For iCol = 1 To spdTAT.DataColCnt
'            spdTAT.Row = iRow
'            spdTAT.Col = iCol
'            xlSheet.Cells(iRow + 1, iCol) = CStr(spdTAT.Text)
'        Next iCol
'    Next iRow
    
    Clipboard.Clear
    
    spdTAT.Col = 2:
    spdTAT.Col2 = spdTAT.MaxCols
    spdTAT.Row = 0
    spdTAT.Row2 = spdTAT.MaxRows
    
    Clipboard.SetText spdTAT.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = "TATOver현황"
    
    Set xlSheet = xlBook.Worksheets(4)
'    xlBook.Worksheets("Sheet1").Name = "전체리스트"
'    Set xlSheet = xlBook.Worksheets("전체리스트")
     
'    For iRow = 0 To spdTOT.DataRowCnt
'        For iCol = 1 To spdTOT.DataColCnt
'            spdTOT.Row = iRow
'            spdTOT.Col = iCol
'            xlSheet.Cells(iRow + 1, iCol) = CStr(spdTOT.Text)
'        Next iCol
'    Next iRow
    
    Clipboard.Clear
    
    spdTOT.Col = 1:
    spdTOT.Col2 = spdTOT.MaxCols
    spdTOT.Row = 0
    spdTOT.Row2 = spdTOT.MaxRows
    
    Clipboard.SetText spdTOT.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = "전체리스트"
    
    xlBook.SaveAs (Filename)
    xlApp.Quit
    
    SaveExcel = True
    
    Exit Function
    
Rst:
    SaveExcel = False

End Function

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
    
    dtpStartDt.Value = Date '- 14
    dtpStopDt.Value = Date
    
    spdSTB.MaxRows = 0
    spdING.MaxRows = 0
    spdTAT.MaxRows = 0
    
    lblDate.Caption = dtpStartDt.Value & "~" & dtpStopDt.Value

    '-- 검사대기현황
    Call GetData_STB
    
    '-- 검사진행현황
    Call GetData_ING

    '-- TAT Over 현황
    Call GetTLAData_TAT

    '-- TAT Over 현황
    Call GetTLAData_TAT_StandBy

    '-- 전체현황
    Call GetData_Tot
    
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
    
    '-- TAT Over 현황(대기현황과 비교)
    Call GetTLAData_TAT_StandBy
    
    '-- 전체현황
    Call GetData_Tot
    
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
    Dim intDay      As Integer
    Dim intSpd      As Integer
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    intDay = 0
    intSpd = 0
    
          'SQL = "Select BCD,PNM,(SEX +'/'+ AGE)as SA,BLDTM, ODYYN as 당일, QUICKYN as 빠른" & vbCr
          SQL = "Select BCD,PNM,(SEX +'/'+ AGE)as SA,ADTM, ODYYN as 당일, QUICKYN as 빠른, PID, WORKNO " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
   ' SQL = SQL & " Where NOT STA IN ('E','O','R','H')" & vbCr
    SQL = SQL & " Where NOT STA IN ('O','R','H')" & vbCr
    SQL = SQL & "   And (ODYYN IS NOT NULL AND ODYYN <> '' OR QUICKYN IS NOT NULL AND QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
   ' SQL = SQL & " Order By BLDTM desc" & vbCr
    SQL = SQL & " Order By ADTM " & vbCr
 
    Call SetSQLData("03.검사대기조회", SQL)
    
    spdSTB.ReDraw = False
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            spdSTB.MaxRows = spdSTB.MaxRows + 1
            intRow = spdSTB.MaxRows
            
            'strDtTmRS = Format(Trim(RS_Ser.Fields("BLDTM")), "MM/DD HH:MM")
            strDtTmRS = Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM")
            
            'Call spdSTB.SetText(1, intRow, Format(Trim(RS_Ser.Fields("BLDTM")), "MM/DD HH:MM"))
            Call spdSTB.SetText(1, intRow, Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM"))
            Call spdSTB.SetText(2, intRow, Trim(RS_Ser.Fields("BCD")) & "")
            Call spdSTB.SetText(3, intRow, Trim(mGetP(RS_Ser.Fields("WORKNO"), 2, "-")))
            Call spdSTB.SetText(4, intRow, Trim(RS_Ser.Fields("PNM")) & "[" & Trim(RS_Ser.Fields("SA")) & "]")
            Call spdSTB.SetText(5, intRow, IIf(Trim(RS_Ser.Fields("당일")) <> "", "V", ""))
            Call spdSTB.SetText(6, intRow, IIf(Trim(RS_Ser.Fields("빠른")) <> "", "V", ""))
            Call spdSTB.SetText(7, intRow, Trim(RS_Ser.Fields("PID")))
            
            If Trim(RS_Ser.Fields("당일") & "") <> "" Then
                intDay = intDay + 1
            End If
            
            If Trim(RS_Ser.Fields("빠른") & "") <> "" Then
                intSpd = intSpd + 1
            End If
            
            
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
    lblCnt.Caption = "당일검체 : " & CStr(intDay) & "      빠른검체 : " & CStr(intSpd)
    'lblStandBy.Caption = spdSTB.MaxRows
    
    spdSTB.ReDraw = True
    
End Sub

Private Sub GetData_Tot()
    Dim intDay      As Integer
    Dim intSpd      As Integer
    Dim strTot      As String
    
    intDay = 0
    intSpd = 0
    
          SQL = "Select '1' as Num, Count(*) as Cnt " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
    SQL = SQL & " Where (ODYYN IS NOT NULL AND ODYYN <> '') " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
    SQL = SQL & " Union All " & vbCr
    SQL = SQL & "Select '2' as Num, Count(*) as Cnt " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
    SQL = SQL & " Where (QUICKYN IS NOT NULL AND QUICKYN <> '') " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
    SQL = SQL & " Order By Num"
 
    Call SetSQLData("07.전체조회", SQL)
    
    spdSTB.ReDraw = False
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            If Trim(RS_Ser.Fields("Num")) = "1" Then
                intDay = Trim(RS_Ser.Fields("Cnt"))
            End If
            If Trim(RS_Ser.Fields("Num")) = "2" Then
                intSpd = Trim(RS_Ser.Fields("Cnt"))
            End If
            RS_Ser.MoveNext
        Loop
        
    End If
    
    RS_Ser.Close
    
    lblTotCnt.Caption = "전체검체 : " & intDay + intSpd & "      당일검체 : " & CStr(intDay) & "      빠른검체 : " & CStr(intSpd)
    
End Sub

Private Sub GetData_Tot_Print()
    Dim intRow      As Long
    Dim strDtTm     As String
    Dim strDtTmRS   As String
    Dim intDay      As Integer
    Dim intSpd      As Integer
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    intDay = 0
    intSpd = 0
    
          SQL = "Select BCD,PNM,(SEX +'/'+ AGE)as SA,ADTM, ODYYN as 당일, QUICKYN as 빠른, PID, WORKNO " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
    SQL = SQL & " Where (ODYYN IS NOT NULL AND ODYYN <> '' OR QUICKYN IS NOT NULL AND QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
    SQL = SQL & " Order By ADTM " & vbCr
    
    Call SetSQLData("08.전체조회", SQL)
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            spdTOT.MaxRows = spdTOT.MaxRows + 1
            intRow = spdTOT.MaxRows
            
            strDtTmRS = Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM")
            
            Call spdTOT.SetText(1, intRow, Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM"))
            Call spdTOT.SetText(2, intRow, Trim(RS_Ser.Fields("BCD")) & "")
            Call spdTOT.SetText(3, intRow, Trim(mGetP(RS_Ser.Fields("WORKNO"), 2, "-")))
            Call spdTOT.SetText(4, intRow, Trim(RS_Ser.Fields("PNM")) & "[" & Trim(RS_Ser.Fields("SA")) & "]")
            Call spdTOT.SetText(5, intRow, IIf(Trim(RS_Ser.Fields("당일")) <> "", "V", ""))
            Call spdTOT.SetText(6, intRow, IIf(Trim(RS_Ser.Fields("빠른")) <> "", "V", ""))
            Call spdTOT.SetText(7, intRow, Trim(RS_Ser.Fields("PID")))
            
            RS_Ser.MoveNext
        Loop
        
        spdTOT.RowHeight(-1) = 10
    End If
    
    RS_Ser.Close
    
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
    Dim strDtTmRS   As String
    
    
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

          SQL = "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '1' as SEQ, a.PID, a.WORKNO, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And (a.ODYYN IS NOT NULL AND a.ODYYN <> '' OR a.QUICKYN IS NOT NULL AND a.QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'O'" & vbCr
    SQL = SQL & "   And a.STA <> 'H'" & vbCr
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
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '2' as SEQ, a.PID, a.WORKNO, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And (a.ODYYN IS NOT NULL AND a.ODYYN <> '' OR a.QUICKYN IS NOT NULL AND a.QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'R'" & vbCr
    SQL = SQL & "   And a.STA <> 'H'" & vbCr
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
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '3' as SEQ, a.PID, a.WORKNO, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And (a.ODYYN IS NOT NULL AND a.ODYYN <> '' OR a.QUICKYN IS NOT NULL AND a.QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'OO'" & vbCr
    SQL = SQL & "   And a.STA <> 'H'" & vbCr
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
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '4' as SEQ, a.PID, a.WORKNO, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b " & vbCr
    SQL = SQL & " Where a.BCD = b.BCD" & vbCr
    SQL = SQL & "   And (a.ODYYN IS NOT NULL AND a.ODYYN <> '' OR a.QUICKYN IS NOT NULL AND a.QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And b.INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "   And b.EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "   And b.STA = 'RR'" & vbCr
    SQL = SQL & "   And a.STA <> 'H'" & vbCr
    SQL = SQL & "   And b.SEQ IN (Select MAX(SEQ)" & vbCr
    SQL = SQL & "                   From " & gDB & "..TB_EVENT " & vbCr
    SQL = SQL & "                  Where EQPCD IN (" & gEQPCD & ")" & vbCr
    SQL = SQL & "                    And INDTM between '" & Format(dtpStartDt.Value, "yyyymmdd") & "000000' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "235959'" & vbCr
    SQL = SQL & "                    And BCD = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    'SQL = SQL & " and a.STA = 'H'"
    SQL = SQL & " Order By a.ADTM, a.BCD ASC,SEQ, b.STA DESC, b.EQPCD ASC " & vbCr
    

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
    
    totO = 0
    totOO = 0
    totR = 0
    totRR = 0
    totH = 0
    totE = 0
    
    intO = 0
    intOO = 0
    intR = 0
    intRR = 0
    intH = 0
    intE = 0
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            If Trim(RS_Ser.Fields("ASTA")) <> "H" Then
                If strBarNo <> Trim(RS_Ser.Fields("BCD")) Then
                    strBarNo = Trim(RS_Ser.Fields("BCD"))
                    
                    'Call Sub_SetColor_STB(strBarNo, &HC0FFC0)    '--필요할지 몰라서..
                    
                    spdING.MaxRows = spdING.MaxRows + 1
                    
                    If Trim(RS_Ser.Fields("ERR")) <> "" Then
                        Call spdING.InsertRows(1, 1)
                        intRow = 1
                    Else
                        intRow = spdING.MaxRows
                    End If
                    
                    Call spdING.SetText(1, intRow, Trim(RS_Ser.Fields("EQPCD")))
                    strDtTmRS = Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM")
                    Call spdING.SetText(2, intRow, strDtTmRS)
                    
                    If DateDiff("n", strDtTmRS, strDtTm) < gLimit Then
                        '배경 - 하얀색
                        spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = vbWhite
                    Else
                        If DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 1) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(0).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 2) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(1).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 3) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(2).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 4) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(3).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 5) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(4).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 6) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(5).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 7) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(6).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 8) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(7).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 9) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(8).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 10) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(9).CaptionBackColor
                        ElseIf DateDiff("n", strDtTmRS, strDtTm) < gLimit + (gLimitS * 11) Then
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(10).CaptionBackColor
                        Else
                            spdING.Row = intRow: spdING.Col = 2: spdING.BackColor = LmColor(11).CaptionBackColor
                        End If
                    End If
                    
                    Call spdING.SetText(3, intRow, Trim(RS_Ser.Fields("BCD")) & "")
                    Call spdING.SetText(4, intRow, Trim(RS_Ser.Fields("PID")))
                    Call spdING.SetText(5, intRow, Trim(mGetP(RS_Ser.Fields("WORKNO"), 2, "-")))
                    Call spdING.SetText(6, intRow, Trim(RS_Ser.Fields("PNM")) & "[" & Trim(RS_Ser.Fields("SA")) & "]")

                    '-- 카운트 올리기
                    If intRR > 0 Then
                        totRR = totRR + 1
                    End If
                    
                    If intOO > 0 Then
                        totOO = totOO + 1
                    End If
                    
                    If intR > 0 Then
                        totR = totR + 1
                    End If
                    
                    If intO > 0 Then
                        totO = totO + 1
                    End If
                    
                    If intE > 0 Then
                        totE = totE + 1
                    End If
                        
                    
                    intO = 0
                    intOO = 0
                    intR = 0
                    intRR = 0
                    intH = 0
                    intE = 0
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
                        If Trim(RS_Ser.Fields("ERR")) & "" = "ERR" Then
                            Call spdING.SetText(intCol, intRow, "에러")
                            spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = &HFF00FF
                            intE = 1
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
            
            RS_Ser.MoveNext
        Loop
        
        
        
        spdING.RowHeight(-1) = 30
    End If
    
    RS_Ser.Close
    
    '-- 카운트 올리기
    If intRR > 0 Then
        totRR = totRR + 1
    End If
    
    If intOO > 0 Then
        totOO = totOO + 1
    End If
    
    If intR > 0 Then
        totR = totR + 1
    End If
    
    If intO > 0 Then
        totO = totO + 1
    End If
    
    If intE > 0 Then
        totE = totE + 1
    End If
    
    lblO.Caption = totO
    lblOO.Caption = totOO
    lblR.Caption = totR
    lblRR.Caption = totRR
    'lblH.Caption = totH
    lblH.Caption = totE


    spdING.ReDraw = True
    
    DoEvents

End Sub

Private Sub GetTLAData_TAT_NEW()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim varTmp      As Variant
    Dim strDtTm     As String
    Dim strNow      As String
    Dim strBldDtTm  As String
    Dim strBarNo    As String
    Dim strPatInfo  As String
    Dim strEqpCd    As String
    Dim strOverTm   As String
    Dim intCnt      As Integer
    Dim strPID      As String
    
'    Dim intCnt      As Integer
'    Dim strGrpCd    As String
'    Dim strCond     As String
'    Dim intO        As Integer
'    Dim intOO       As Integer
'    Dim intR        As Integer
'    Dim intRR       As Integer
'    Dim intH        As Integer

    On Error Resume Next
    
    spdTAT.ReDraw = False
    spdTAT.MaxRows = 0
    intCnt = 0
        
    For intRow = 1 To spdSTB.MaxRows
        Call spdSTB.GetText(1, intRow, varTmp): strBldDtTm = varTmp
        Call spdSTB.GetText(2, intRow, varTmp): strBarNo = varTmp
        Call spdSTB.GetText(3, intRow, varTmp): strPatInfo = varTmp
        Call spdSTB.GetText(6, intRow, varTmp): strPID = varTmp
        
        strNow = Format(Now(), "yyyy-MM-dd hh:nn:ss")
        strOverTm = DateDiff("n", strBldDtTm, strNow)

            If strOverTm > 120 Then
                spdTAT.MaxRows = spdTAT.MaxRows + 1
                intCnt = intCnt + 1
                Call spdTAT.SetText(2, intCnt, strBarNo)
                Call spdTAT.SetText(3, intCnt, strPatInfo)
                Call spdTAT.SetText(4, intCnt, strPID)
                Call spdTAT.SetText(5, intCnt, strOverTm & " (+" & strOverTm - "120" & "m)")
            End If
        
    Next

    spdTAT.RowHeight(-1) = 30
    lblTAT.Caption = intCnt
    
    
    spdTAT.ReDraw = True
    
    

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
    Dim strPID      As String
    Dim intCnt      As Integer
    Dim strWN       As String
    
'    Dim intCnt      As Integer
'    Dim strGrpCd    As String
'    Dim strCond     As String
'    Dim intO        As Integer
'    Dim intOO       As Integer
'    Dim intR        As Integer
'    Dim intRR       As Integer
'    Dim intH        As Integer

    On Error Resume Next
    
    spdTAT.ReDraw = False
    spdTAT.MaxRows = 0
    intCnt = 0
        
    For intRow = 1 To spdING.MaxRows
        Call spdING.GetText(2, intRow, varTmp): strBldDtTm = varTmp
        Call spdING.GetText(3, intRow, varTmp): strBarNo = varTmp
        Call spdING.GetText(4, intRow, varTmp): strPID = varTmp
        Call spdING.GetText(5, intRow, varTmp): strWN = varTmp
        Call spdING.GetText(6, intRow, varTmp): strPatInfo = varTmp
        
        For intCol = 7 To 10
            Call spdING.GetText(intCol, intRow, varTmp): strEqpCd = varTmp
            If strEqpCd <> "" Then
                Call spdING.GetText(intCol, 0, varTmp): strEqpCd = varTmp
                
'                      SQL = "Select INDTM, (Select DISTINCT PID From " & gDB & "..TB_BCD Where  BCD = '" & strBarNo & "') as PID " & vbCr
'                SQL = SQL & "  From " & gDB & "..TB_EVENT " & vbCr
'                SQL = SQL & " Where BCD = '" & strBarNo & "'" & vbCr
'                SQL = SQL & "   And GRPCD = '" & strEqpCd & "'" & vbCr
'                SQL = SQL & "   And STA IN ('O','R') " & vbCr
'                SQL = SQL & "   And SEQ = (Select MAX(SEQ) FROM " & gDB & "..TB_EVENT "
'                SQL = SQL & "               Where BCD = '" & strBarNo & "'" & vbCr
'                SQL = SQL & "                 And GRPCD = '" & strEqpCd & "'" & vbCr
'                SQL = SQL & "                 And STA IN ('O','R') " & vbCr
'                SQL = SQL & "               Group By BCD) " & vbCr
                
'''                      SQL = "Select a.ADTM, (Select DISTINCT PID From " & gDB & "..TB_BCD Where  BCD = '" & strBarNo & "') as PID " & vbCr
'''                SQL = SQL & "  From " & gDB & "..TB_BCD a, " & gDB & "..TB_EVENT b" & vbCr
'''                SQL = SQL & " Where a.BCD = b.BCD "
'''                SQL = SQL & "   And b.BCD = '" & strBarNo & "'" & vbCr
'''                SQL = SQL & "   And b.GRPCD = '" & strEqpCd & "'" & vbCr
'''                SQL = SQL & "   And b.STA IN ('O','R') " & vbCr
'''                SQL = SQL & "   And b.SEQ = (Select MAX(SEQ) FROM " & gDB & "..TB_EVENT "
'''                SQL = SQL & "               Where BCD = '" & strBarNo & "'" & vbCr
'''                SQL = SQL & "                 And GRPCD = '" & strEqpCd & "'" & vbCr
'''                SQL = SQL & "                 And STA IN ('O','R') " & vbCr
'''                SQL = SQL & "               Group By BCD) " & vbCr
'''
'''                Call SetSQLData("05.TAT조회", SQL)
'''
'''                Set RS_Ser = Cn_Ser.Execute(SQL)
'''                If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
'''                    Do Until RS_Ser.EOF
                        strDtTm = CStr(Format(Now, "yyyy-mm-dd hh:mm:ss"))
                        'strDtTm = Format(RS_Ser.Fields("ADTM"), "####-##-## ##:##:##")
                        strDtTm = Format(strDtTm, "MM/DD HH:MM")
                        strOverTm = DateDiff("n", strBldDtTm, strDtTm)
                        If intCol = 6 Then  'ARC
                            If strOverTm > gTatARC Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo & "")
                                Call spdTAT.SetText(3, intCnt, strWN)
                                Call spdTAT.SetText(4, intCnt, strPatInfo)
                                Call spdTAT.SetText(5, intCnt, strPID)
                                Call spdTAT.SetText(6, intCnt, "ARC")
                                Call spdTAT.SetText(7, intCnt, strOverTm & " (+" & strOverTm - gTatARC & "m)")
                            End If
                        ElseIf intCol = 7 Then  'AU
                            If strOverTm > gTatAU Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo & "")
                                Call spdTAT.SetText(3, intCnt, strWN)
                                Call spdTAT.SetText(4, intCnt, strPatInfo)
                                Call spdTAT.SetText(5, intCnt, strPID)
                                Call spdTAT.SetText(6, intCnt, "AU")
                                Call spdTAT.SetText(7, intCnt, strOverTm & " (+" & strOverTm - gTatAU & "m)")
                            End If
                        ElseIf intCol = 8 Then  'C16000
                            If strOverTm > gTatTAT Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo & "")
                                Call spdTAT.SetText(3, intCnt, strWN)
                                Call spdTAT.SetText(4, intCnt, strPatInfo)
                                Call spdTAT.SetText(5, intCnt, strPID)
                                Call spdTAT.SetText(6, intCnt, "C16000")
                                Call spdTAT.SetText(7, intCnt, strOverTm & " (+" & strOverTm - gTatCOB & "m)")
                            End If
                        ElseIf intCol = 9 Then  'COB
                            If strOverTm > gTatCOB Then
                                spdTAT.MaxRows = spdTAT.MaxRows + 1
                                intCnt = intCnt + 1
                                Call spdTAT.SetText(2, intCnt, strBarNo & "")
                                Call spdTAT.SetText(3, intCnt, strWN)
                                Call spdTAT.SetText(4, intCnt, strPatInfo)
                                Call spdTAT.SetText(5, intCnt, strPID)
                                Call spdTAT.SetText(6, intCnt, "COB")
                                Call spdTAT.SetText(7, intCnt, strOverTm & " (+" & strOverTm - gTatCOB & "m)")
                            End If
                        End If
                        

'''                        RS_Ser.MoveNext
'''                    Loop
'''                End If
'''                RS_Ser.Close
            End If
        Next
    Next

    spdTAT.RowHeight(-1) = 30
    lblTAT.Caption = intCnt
    
    
    spdTAT.ReDraw = True

End Sub

Private Sub GetTLAData_TAT_StandBy()
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
    Dim strPID      As String
    Dim intTot      As Integer
    Dim strWN       As String
    

    On Error Resume Next
    
    spdTAT.ReDraw = False
    'spdTAT.MaxRows = 0
    intCnt = 0
        
    For intRow = 1 To spdSTB.MaxRows
        Call spdSTB.GetText(1, intRow, varTmp): strBldDtTm = varTmp
        Call spdSTB.GetText(2, intRow, varTmp): strBarNo = varTmp
        Call spdSTB.GetText(3, intRow, varTmp): strWN = varTmp
        Call spdSTB.GetText(4, intRow, varTmp): strPatInfo = varTmp
        Call spdSTB.GetText(7, intRow, varTmp): strPID = varTmp
        
        'strDtTm = Format(RS_Ser.Fields("INDTM"), "####-##-## ##:##:##")
        strDtTm = Format(Now, "MM/DD HH:MM")
        strOverTm = DateDiff("n", strBldDtTm, strDtTm)
        
        If strOverTm > gTatTAT Then
            spdTAT.MaxRows = spdTAT.MaxRows + 1
            
            Call spdTAT.SetText(2, spdTAT.MaxRows, strBarNo & "")
            Call spdTAT.SetText(3, spdTAT.MaxRows, strWN)
            Call spdTAT.SetText(4, spdTAT.MaxRows, strPatInfo)
            Call spdTAT.SetText(5, spdTAT.MaxRows, strPID)
            Call spdTAT.SetText(6, spdTAT.MaxRows, "TAT")
            Call spdTAT.SetText(7, spdTAT.MaxRows, strOverTm & " (+" & strOverTm - gTatTAT & "m)")
            intCnt = intCnt + 1
        End If
    Next

    spdTAT.RowHeight(-1) = 30
    intTot = lblTAT.Caption
    lblTAT.Caption = intTot + intCnt
    
    
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
    spdTOT.MaxRows = 0
    
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
    lblCnt.Caption = ""
    lblTotCnt.Caption = ""
    
    fraSet.Height = 5385
    
End Sub

Private Sub GetIni()
    Dim DB_Tmp  As String * 100
    
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
    Call GetPrivateProfileString("TLA", "TATTAT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtTATTAT.Text = Trim(txtTemp.Text)
    gTatTAT = txtTATTAT.Text
    
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
    
    '색설정
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV1", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV1 = Trim(txtTemp.Text)
    If BGColor.LV1 <> "" Then
        LmColor(0).CaptionBackColor = CCur(BGColor.LV1)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV2", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV2 = Trim(txtTemp.Text)
    If BGColor.LV2 <> "" Then
        LmColor(1).CaptionBackColor = CCur(BGColor.LV2)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV3", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV3 = Trim(txtTemp.Text)
    If BGColor.LV3 <> "" Then
        LmColor(2).CaptionBackColor = CCur(BGColor.LV3)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV4", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV4 = Trim(txtTemp.Text)
    If BGColor.LV4 <> "" Then
        LmColor(3).CaptionBackColor = CCur(BGColor.LV4)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV5", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV5 = Trim(txtTemp.Text)
    If BGColor.LV5 <> "" Then
        LmColor(4).CaptionBackColor = CCur(BGColor.LV5)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV6", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV6 = Trim(txtTemp.Text)
    If BGColor.LV6 <> "" Then
        LmColor(5).CaptionBackColor = CCur(BGColor.LV6)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV7", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV7 = Trim(txtTemp.Text)
    If BGColor.LV7 <> "" Then
        LmColor(6).CaptionBackColor = CCur(BGColor.LV7)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV8", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV8 = Trim(txtTemp.Text)
    If BGColor.LV8 <> "" Then
        LmColor(7).CaptionBackColor = CCur(BGColor.LV8)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV9", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV9 = Trim(txtTemp.Text)
    If BGColor.LV9 <> "" Then
        LmColor(8).CaptionBackColor = CCur(BGColor.LV9)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV10", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV10 = Trim(txtTemp.Text)
    If BGColor.LV10 <> "" Then
        LmColor(9).CaptionBackColor = CCur(BGColor.LV10)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV11", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV11 = Trim(txtTemp.Text)
    If BGColor.LV11 <> "" Then
        LmColor(10).CaptionBackColor = CCur(BGColor.LV11)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LV12", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LV12 = Trim(txtTemp.Text)
    If BGColor.LV12 <> "" Then
        LmColor(11).CaptionBackColor = CCur(BGColor.LV12)
    End If
    
    
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

Private Sub Sub_SetColor_STB(pBCD As String, BackColor As OLE_COLOR)

    Dim i As Integer
    Dim BCD As String
    

    For i = 1 To spdSTB.MaxRows
        Call spdSTB.GetText(2, i, BCD)
        If BCD = pBCD Then
            spdSTB.Row = i: spdSTB.Col = 2: spdSTB.BackColor = BackColor
            Exit For
        End If

    Next
                

End Sub
