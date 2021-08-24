VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "GTCotrol.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTLAMonitor 
   Appearance      =   0  '평면
   BackColor       =   &H00F8E4D8&
   Caption         =   "[TLA 모니터]"
   ClientHeight    =   12435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   24495
   Icon            =   "frmTLAMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12435
   ScaleWidth      =   24495
   StartUpPosition =   2  '화면 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame fraTAT 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   15630
      TabIndex        =   8
      Top             =   180
      Width           =   6675
      Begin HSCotrol.CButton cmdExit 
         Height          =   570
         Left            =   5370
         TabIndex        =   17
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
         Left            =   1410
         TabIndex        =   18
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
         Left            =   2730
         TabIndex        =   27
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
         Left            =   4050
         TabIndex        =   84
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
         TabIndex        =   88
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
         Left            =   300
         TabIndex        =   56
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
         Left            =   5400
         TabIndex        =   55
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         Height          =   885
         Left            =   5130
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
         Left            =   120
         TabIndex        =   19
         Top             =   690
         Width           =   1245
      End
      Begin VB.Image Image3 
         Height          =   225
         Left            =   120
         Picture         =   "frmTLAMonitor.frx":144A
         Top             =   240
         Width           =   150
      End
      Begin VB.Label lblName 
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
         Index           =   2
         Left            =   420
         TabIndex        =   16
         Top             =   210
         Width           =   2745
      End
   End
   Begin VB.Frame fraING 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5880
      TabIndex        =   6
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   89
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
         TabIndex        =   53
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1350
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   120
         Picture         =   "frmTLAMonitor.frx":1834
         Top             =   210
         Width           =   150
      End
      Begin VB.Label lblName 
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
         Index           =   1
         Left            =   420
         TabIndex        =   15
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
      Height          =   8385
      Left            =   9555
      TabIndex        =   20
      Top             =   2580
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox txtC16TAT 
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
         Left            =   2430
         TabIndex        =   85
         Text            =   "120"
         Top             =   2850
         Width           =   1395
      End
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   0
         Left            =   1290
         TabIndex        =   71
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   58
         Top             =   5820
         Width           =   1005
         _ExtentX        =   1773
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
         TabIndex        =   50
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
         TabIndex        =   48
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
         Left            =   2430
         TabIndex        =   42
         Text            =   "120"
         Top             =   2460
         Width           =   1395
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
         Left            =   2430
         TabIndex        =   40
         Text            =   "90"
         Top             =   2070
         Width           =   1395
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
         Left            =   2430
         TabIndex        =   37
         Text            =   "120"
         Top             =   1680
         Width           =   1395
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
         TabIndex        =   32
         Text            =   "30"
         Top             =   150
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   435
         Left            =   3390
         TabIndex        =   28
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
         Format          =   70778881
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   435
         Left            =   1260
         TabIndex        =   29
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
         Format          =   70778881
         CurrentDate     =   40248
      End
      Begin HSCotrol.CButton cmdApply 
         Height          =   480
         Left            =   2670
         TabIndex        =   30
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
         TabIndex        =   31
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
         Left            =   270
         TabIndex        =   59
         Top             =   6210
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   270
         TabIndex        =   60
         Top             =   6600
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   270
         TabIndex        =   61
         Top             =   6990
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   270
         TabIndex        =   62
         Top             =   7380
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   270
         TabIndex        =   63
         Top             =   7770
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   64
         Top             =   5820
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   65
         Top             =   6210
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   66
         Top             =   6600
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   67
         Top             =   6990
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   68
         Top             =   7380
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2070
         TabIndex        =   69
         Top             =   7770
         Width           =   1005
         _ExtentX        =   1773
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
         TabIndex        =   70
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
         Left            =   1290
         TabIndex        =   72
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
         Left            =   1290
         TabIndex        =   73
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
         Left            =   1290
         TabIndex        =   74
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
         Left            =   1290
         TabIndex        =   75
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
         Left            =   1290
         TabIndex        =   76
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
         Left            =   3090
         TabIndex        =   77
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
         Left            =   3090
         TabIndex        =   78
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
         Left            =   3090
         TabIndex        =   79
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
         Left            =   3090
         TabIndex        =   80
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
         Left            =   3090
         TabIndex        =   81
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
         Left            =   3090
         TabIndex        =   82
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   12
         Left            =   3870
         TabIndex        =   91
         Top             =   5820
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         CaptionBackColor=   12648447
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   12
         Left            =   4890
         TabIndex        =   92
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   13
         Left            =   3870
         TabIndex        =   93
         Top             =   6210
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         CaptionBackColor=   65535
         Caption         =   "재검"
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   13
         Left            =   4890
         TabIndex        =   94
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   14
         Left            =   3870
         TabIndex        =   95
         Top             =   6600
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         CaptionBackColor=   12648384
         Caption         =   "결과"
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   14
         Left            =   4890
         TabIndex        =   96
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   15
         Left            =   3870
         TabIndex        =   97
         Top             =   6990
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         CaptionBackColor=   65280
         Caption         =   "재결"
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   15
         Left            =   4890
         TabIndex        =   98
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
      Begin HSCotrol.CaptionBar LmColor 
         Height          =   375
         Index           =   16
         Left            =   3870
         TabIndex        =   99
         Top             =   7380
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         CaptionBackColor=   16711935
         Caption         =   "에러"
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
      Begin HSCotrol.CButton cmdColorSet 
         Height          =   390
         Index           =   16
         Left            =   4890
         TabIndex        =   100
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "C16000"
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
         TabIndex        =   87
         Top             =   2970
         Width           =   855
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
         TabIndex        =   86
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   41
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
   Begin VB.Frame fraSTB 
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   180
      TabIndex        =   7
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
         TabIndex        =   83
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
         TabIndex        =   57
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label lblStandBy 
         BackStyle       =   0  '투명
         Caption         =   "대기 123 검체"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   39.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1005
         Left            =   30
         TabIndex        =   54
         Top             =   840
         Width           =   6435
      End
      Begin VB.Label lblName 
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
         TabIndex        =   9
         Top             =   150
         Width           =   2745
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   120
         Picture         =   "frmTLAMonitor.frx":1C1E
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
      Caption         =   "Hidden Control"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2925
      Left            =   630
      TabIndex        =   3
      Top             =   8220
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CheckBox chkS 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1950
         TabIndex        =   22
         Top             =   930
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Timer tmrMonitor 
         Left            =   2850
         Top             =   900
      End
      Begin VB.Timer tmrRefresh 
         Left            =   3270
         Top             =   900
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
         Height          =   480
         Left            =   270
         TabIndex        =   5
         Text            =   "123456789012"
         Top             =   330
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   270
         TabIndex        =   4
         Top             =   900
         Width           =   1425
      End
      Begin FPSpreadADO.fpSpread spdGrEq 
         Height          =   525
         Left            =   270
         TabIndex        =   21
         Top             =   1380
         Width           =   3495
         _Version        =   524288
         _ExtentX        =   6165
         _ExtentY        =   926
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
         SpreadDesigner  =   "frmTLAMonitor.frx":2008
      End
      Begin FPSpreadADO.fpSpread spdTOT 
         Height          =   615
         Left            =   270
         TabIndex        =   90
         Top             =   2010
         Width           =   3480
         _Version        =   524288
         _ExtentX        =   6138
         _ExtentY        =   1085
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
         SpreadDesigner  =   "frmTLAMonitor.frx":2460
         UserResize      =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3240
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin FPSpreadADO.fpSpread spdSTB 
      Height          =   8895
      Left            =   180
      TabIndex        =   0
      Top             =   2550
      Width           =   5400
      _Version        =   524288
      _ExtentX        =   9525
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
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
      SpreadDesigner  =   "frmTLAMonitor.frx":2B76
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
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
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
      MaxCols         =   8
      MaxRows         =   499
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmTLAMonitor.frx":32F5
      UserResize      =   0
   End
   Begin FPSpreadADO.fpSpread spdTAT 
      Height          =   8895
      Left            =   15630
      TabIndex        =   2
      Top             =   2550
      Width           =   6750
      _Version        =   524288
      _ExtentX        =   11906
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
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
      SpreadDesigner  =   "frmTLAMonitor.frx":3D81
      UserResize      =   0
   End
End
Attribute VB_Name = "frmTLAMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmTLAMonitor.frm
'   작성자  : 오세원
'   내  용  : TLA 검사진행 상태 모니터링
'   작성일  :
'   버  전  : 1.0.0
'   병  원  : 분당서울대병원 진단검사의학과 자동화검사실
'-----------------------------------------------------------------------------'

Option Explicit

Private lngTime As Long


Private Sub cmdApply_Click()
    
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
        
End Sub

Private Sub cmdColor_Click()
    
    If fraSet.Height = 8385 Then
        fraSet.Height = 5385
    Else
        fraSet.Height = 8385
    End If
    
End Sub

Private Sub cmdColorSet_Click(Index As Integer)
    Dim LetColor
    Dim GetColor
    
    '선택하기전 색을 갖고 있는다.
    LetColor = LmColor(Index).BackColor
    
    CommonDialog1.ShowColor
    
    '선택한 색이 없다면
    If CommonDialog1.Color = 0 Then
        LmColor(Index).CaptionBackColor = LetColor
    '............있다면
    Else
        GetColor = CommonDialog1.Color
        LmColor(Index).CaptionBackColor = GetColor
    End If
    
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
            
        Case 12:    Call WritePrivateProfileString("COLOR", "LVO", CStr(GetColor), App.Path & "\TLA.ini")
        Case 13:    Call WritePrivateProfileString("COLOR", "LVOO", CStr(GetColor), App.Path & "\TLA.ini")
        Case 14:    Call WritePrivateProfileString("COLOR", "LVR", CStr(GetColor), App.Path & "\TLA.ini")
        Case 15:    Call WritePrivateProfileString("COLOR", "LVRR", CStr(GetColor), App.Path & "\TLA.ini")
        Case 16:    Call WritePrivateProfileString("COLOR", "LVH", CStr(GetColor), App.Path & "\TLA.ini")
    End Select
    
End Sub

Private Sub cmdExcel_Click()
    Dim sFileName As String
    
On Error GoTo Rst
    
    Call GetData_Tot_Print

    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    sFileName = CommonDialog1.Filename
    
    If SaveExcel(sFileName) = True Then
        MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
    Else
        MsgBox "엑셀 저장실패", vbOKOnly + vbInformation, Me.Caption
    End If
    
    Exit Sub
    
Rst:
    '
    
End Sub

Private Function SaveExcel(Filename As String) As Boolean

On Error GoTo Rst
    
    SaveExcel = False
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    '=========================================
    Set xlSheet = xlBook.Worksheets(1)
    Clipboard.Clear
    spdSTB.Col = 1:     spdSTB.Col2 = spdSTB.MaxCols
    spdSTB.Row = 0:     spdSTB.Row2 = spdSTB.MaxRows
    Clipboard.SetText spdSTB.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = lblName(0).Caption
    Clipboard.Clear
    '=========================================
    
    '=========================================
    Set xlSheet = xlBook.Worksheets(2)
    spdING.Col = 2:     spdING.Col2 = spdING.MaxCols
    spdING.Row = 0:     spdING.Row2 = spdING.MaxRows
    Clipboard.SetText spdING.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = lblName(1).Caption
    Clipboard.Clear
    '=========================================
    
    '=========================================
    Set xlSheet = xlBook.Worksheets(3)
    spdTAT.Col = 2:     spdTAT.Col2 = spdTAT.MaxCols
    spdTAT.Row = 0:     spdTAT.Row2 = spdTAT.MaxRows
    Clipboard.SetText spdTAT.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = lblName(2).Caption
    Clipboard.Clear
    '=========================================
    
    '=========================================
    Set xlSheet = xlBook.Worksheets(4)
    spdTOT.Col = 1:     spdTOT.Col2 = spdTOT.MaxCols
    spdTOT.Row = 0:     spdTOT.Row2 = spdTOT.MaxRows
    Clipboard.SetText spdTOT.Clip
    xlSheet.Paste
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Name = "전체리스트"
    Clipboard.Clear
    '=========================================
    
    xlBook.SaveAs (Filename)
    xlApp.Quit
    
    SaveExcel = True
    
    Exit Function
    
Rst:
    SaveExcel = False

End Function


Private Sub cmdExit_Click()
    
    Unload Me

End Sub

Private Sub cmdRefresh_Click()
    
    dtpStartDt.Value = Date
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

    '-- TAT Over 현황(대기현황과 비교)
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
    
    '-- 폼 컨트롤 초기화
    Call FrmInitial
    
    '-- 설정파일(TLA.ini)
    Call GetIni

    '-- Refresh 타이머 시작
    tmrRefresh.Interval = lblSec * 1000
    tmrRefresh.Enabled = True

    '-- DB 연결
    If Connect_PRServer Then
        lblStatus.Caption = "TLA 서버 연결 성공"
        tmrMonitor.Interval = 500
        tmrMonitor.Enabled = True
    Else
        lblStatus.Caption = "TLA 서버 연결 실패"
        Exit Sub
    End If

    '-- 01.그룹리스트조회
    Call GetGRPList

    '-- 01-2.작업리스트조회
    Call GetWRKList
    
    '-- 02.장비리스트조회
    Call GetEQPList

    '-- 03.검사대기현황
    Call GetData_STB

    '-- 04.검사진행현황
    Call GetData_ING

    '-- 05.TAT Over 현황
    Call GetTLAData_TAT

    '-- 06.TAT Over 현황(대기현황과 비교)
    Call GetTLAData_TAT_StandBy

    '-- 07.전체현황
    Call GetData_Tot
    
End Sub

'-- 02.장비리스트조회
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

'-- 01.그룹리스트조회
Private Sub GetGRPList()
    Dim intCol As Long
    
    intCol = 0
    
          SQL = "Select DISTINCT GRPCD" & vbCr
    SQL = SQL & "  From " & gDB & "..TB_EQP" & vbCr
    SQL = SQL & " Where LISCD < 1000" & vbCr
    SQL = SQL & " Order By GRPCD " & vbCr
 
 
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    Call SetSQLData("01.그룹리스트조회", SQL)
    
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        'With spdING
            '.MaxCols = colHeader + RS_Ser.RecordCount
            Do Until RS_Ser.EOF
            '    intCol = intCol + 1
            '    .Col = colHeader + intCol
            '    .Row = -1
            '    .CellType = CellTypeStaticText
            '    .TypeEditCharSet = TypeEditCharSetASCII
            '    .TypeEditCharCase = TypeEditCharCaseSetNone
            '    .TypeHAlign = TypeHAlignCenter
            '    .TypeVAlign = TypeVAlignCenter
            '    .ColWidth(colHeader + intCol) = gWIDTH
            '    Call spdING.SetText(colHeader + intCol, 0, Trim(RS_Ser.Fields("GRPCD")))
                gGRPCD = gGRPCD & "'" & Trim(RS_Ser.Fields("GRPCD")) & "',"
                RS_Ser.MoveNext
            Loop
            RS_Ser.Close
        'End With
        
        'If gGRPCD <> "" Then
        '    gGRPCD = Mid(gGRPCD, 1, Len(gGRPCD) - 1)
        'End If
    End If
    
End Sub

'-- 01-2.작업리스트조회
Private Sub GetWRKList()
    Dim intCol As Long
    
    intCol = 0
    
          SQL = "Select DISTINCT EQPCD " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_EQP" & vbCr
    SQL = SQL & " Where LISCD < 1000" & vbCr
    SQL = SQL & "   And EQPCD <> 'AU' "
    SQL = SQL & " Order By EQPCD " & vbCr
 
 
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    Call SetSQLData("01-2.작업리스트조회", SQL)
    
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
                Call spdING.SetText(colHeader + intCol, 0, Trim(RS_Ser.Fields("EQPCD")))
                RS_Ser.MoveNext
            Loop
            RS_Ser.Close
        End With
        
    End If
    
End Sub


'-- 03.검사대기현황
Private Sub GetData_STB()
    Dim intRow      As Long
    Dim strDtTm     As String
    Dim strDtTmRS   As String
    Dim intDay      As Integer
    Dim intSpd      As Integer
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    intDay = 0
    intSpd = 0
    
          SQL = "Select BCD, PNM, (SEX +'/'+ AGE) as SA, ADTM, ODYYN as 당일, QUICKYN as 빠른, PID, WORKNO " & vbCr
    SQL = SQL & "  From " & gDB & "..TB_BCD" & vbCr
    SQL = SQL & " Where NOT STA IN ('O','R','H')" & vbCr
    SQL = SQL & "   And (ODYYN IS NOT NULL AND ODYYN <> '' OR QUICKYN IS NOT NULL AND QUICKYN <> '' ) " & vbCr
    SQL = SQL & "   And INDT between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
    SQL = SQL & " Order By ADTM " & vbCr
    Set RS_Ser = Cn_Ser.Execute(SQL)
 
    Call SetSQLData("03.검사대기조회", SQL)
    spdSTB.ReDraw = False
    
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            spdSTB.MaxRows = spdSTB.MaxRows + 1
            intRow = spdSTB.MaxRows
            strDtTmRS = Format(Trim(RS_Ser.Fields("ADTM")), "MM/DD HH:MM")
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
        spdSTB.RowHeight(-1) = 20 '30
    End If
    
    RS_Ser.Close
    
    lblStandBy.Caption = "대기 " & spdSTB.MaxRows & " 검체"
    lblCnt.Caption = "당일검체 : " & CStr(intDay) & "      빠른검체 : " & CStr(intSpd)
    
    spdSTB.ReDraw = True
    
End Sub

'-- 07.전체현황
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
    
          SQL = "Select BCD,PNM,(SEX +'/'+ AGE) as SA,ADTM, ODYYN as 당일, QUICKYN as 빠른, PID, WORKNO " & vbCr
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

'-- 04.검사진행현황
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
    strGrpCd = ""
    strBarNo = ""
    
    lblRR.Caption = 0
    lblOO.Caption = 0
    lblR.Caption = 0
    lblO.Caption = 0
    lblH.Caption = 0
    
    totO = 0: totOO = 0: totR = 0: totRR = 0: totH = 0: totE = 0
    intO = 0: intOO = 0: intR = 0: intRR = 0: intH = 0: intE = 0
    
    If chkO.Value = "1" Then strCond = strCond & "'O',"
    If chkOO.Value = "1" Then strCond = strCond & "'OO',"
    If chkR.Value = "1" Then strCond = strCond & "'R',"
    If chkRR.Value = "1" Then strCond = strCond & "'RR',"
    If chkH.Value = "1" Then strCond = strCond & "'H',"
    If chkS.Value = "1" Then strCond = strCond & "'S',"
    
    If strCond <> "" Then
        strCond = Mid(strCond, 1, Len(strCond) - 1)
    End If
    
          SQL = "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '1' as SEQ, a.PID, a.WORKNO, ODYYN as 당일, QUICKYN as 빠른, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
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
    SQL = SQL & "                    And BCD   = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '2' as SEQ, a.PID, a.WORKNO, ODYYN as 당일, QUICKYN as 빠른, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
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
    SQL = SQL & "                    And BCD   = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '3' as SEQ, a.PID, a.WORKNO, ODYYN as 당일, QUICKYN as 빠른, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
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
    SQL = SQL & "                    And BCD   = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    SQL = SQL & " Union All "
    SQL = SQL & "Select DISTINCT a.BCD,a.PNM,(a.SEX +'/'+ a.AGE) as SA,a.ADTM,b.EQPCD,b.STA, a.STA as ASTA, '4' as SEQ, a.PID, a.WORKNO, ODYYN as 당일, QUICKYN as 빠른, (Select top 1 'ERR' From " & gDB & "..TB_Event Where BCD = a.BCD and AREA = 'Er') as ERR" & vbCr
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
    SQL = SQL & "                    And BCD   = b.BCD " & vbCr
    SQL = SQL & "                    And GRPCD = b.GRPCD " & vbCr
    SQL = SQL & "                    And EQPCD = b.EQPCD " & vbCr
    SQL = SQL & "                  Group By BCD, EQPCD)" & vbCr
    SQL = SQL & " Order By a.ADTM, a.BCD ASC,SEQ, b.STA DESC, b.EQPCD ASC " & vbCr

    Call SetSQLData("04.검사진행조회", SQL)

    spdING.ReDraw = False
    
    strDtTm = Format(Now, "MM/DD HH:MM")
    
    Set RS_Ser = Cn_Ser.Execute(SQL)
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        Do Until RS_Ser.EOF
            If Trim(RS_Ser.Fields("ASTA")) <> "H" Then
                If strBarNo <> Trim(RS_Ser.Fields("BCD")) Then
                    strBarNo = Trim(RS_Ser.Fields("BCD"))
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
                    Call spdING.SetText(7, intRow, Trim(RS_Ser.Fields("당일")))
                    Call spdING.SetText(8, intRow, Trim(RS_Ser.Fields("빠른")))

                    '-- 카운트 올리기
                    If intRR > 0 Then totRR = totRR + 1
                    If intOO > 0 Then totOO = totOO + 1
                    If intR > 0 Then totR = totR + 1
                    If intO > 0 Then totO = totO + 1
                    If intE > 0 Then totE = totE + 1
                    
                    '-- 초기화
                    intO = 0: intOO = 0: intR = 0: intRR = 0: intH = 0: intE = 0
                End If
                
                '-- 그룹명 찾기
                strGrpCd = ""
                For intCnt = 1 To spdGrEq.MaxRows
                    spdGrEq.Row = intCnt
                    spdGrEq.Col = 1
                    If Trim(spdGrEq.Text) = Trim(RS_Ser.Fields("EQPCD")) Then
                        'spdGrEq.Col = 2
                        spdGrEq.Col = 1
                        strGrpCd = Trim(spdGrEq.Text)
                        Exit For
                    End If
                Next
                
                For intCol = colHeader To spdING.MaxCols
                    spdING.Col = intCol
                    spdING.Row = 0
                    'If Trim(spdING.Text) = strGrpCd Then
                    If Trim(spdING.Text) = strGrpCd Then
                        If Trim(RS_Ser.Fields("ERR")) & "" = "ERR" Then
                            Call spdING.SetText(intCol, intRow, "에러")
                            spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = BGColor.LVH '&HFF00FF

                            intE = 1
                        Else
                            Select Case Trim(RS_Ser.Fields("STA"))
                            Case "O":
                                        Call spdING.SetText(intCol, intRow, "검사중")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = BGColor.LVO '&HC0FFFF
                                        intO = 1
                            Case "OO":
                                        Call spdING.SetText(intCol, intRow, "재검")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = BGColor.LVOO '&HFFFF&
                                        intOO = 1
                            Case "R":
                                        Call spdING.SetText(intCol, intRow, "결과")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = BGColor.LVR '&HC0FFC0
                                        intR = 1
                            Case "RR":
                                        Call spdING.SetText(intCol, intRow, "재결")
                                        spdING.Row = intRow: spdING.Col = intCol: spdING.BackColor = BGColor.LVRR '&HFF00&
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
        spdING.RowHeight(-1) = 20 '30
    End If
    
    RS_Ser.Close
    
    '-- 마지막건수 연산
    If intRR > 0 Then totRR = totRR + 1
    If intOO > 0 Then totOO = totOO + 1
    If intR > 0 Then totR = totR + 1
    If intO > 0 Then totO = totO + 1
    If intE > 0 Then totE = totE + 1
    
    lblO.Caption = totO
    lblOO.Caption = totOO
    lblR.Caption = totR
    lblRR.Caption = totRR
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

    spdTAT.RowHeight(-1) = 20 '30
    lblTAT.Caption = intCnt
    
    spdTAT.ReDraw = True
    
End Sub

'-- 05.TAT Over 현황
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
        
        For intCol = colHeader + 1 To colHeader + 4
            Call spdING.GetText(intCol, intRow, varTmp): strEqpCd = varTmp
            If strEqpCd <> "" Then
                Call spdING.GetText(intCol, 0, varTmp): strEqpCd = varTmp
                strDtTm = CStr(Format(Now, "yyyy-mm-dd hh:mm:ss"))
                strDtTm = Format(strDtTm, "MM/DD HH:MM")
                strOverTm = DateDiff("n", strBldDtTm, strDtTm)
                If intCol = 9 Then          'ARCHITECT
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
                ElseIf intCol = 10 Then     'AU5800
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
                ElseIf intCol = 11 Then     'C16000
                    If strOverTm > gTatC16 Then
                        spdTAT.MaxRows = spdTAT.MaxRows + 1
                        intCnt = intCnt + 1
                        Call spdTAT.SetText(2, intCnt, strBarNo & "")
                        Call spdTAT.SetText(3, intCnt, strWN)
                        Call spdTAT.SetText(4, intCnt, strPatInfo)
                        Call spdTAT.SetText(5, intCnt, strPID)
                        Call spdTAT.SetText(6, intCnt, "C16000")
                        Call spdTAT.SetText(7, intCnt, strOverTm & " (+" & strOverTm - gTatC16 & "m)")
                    End If
                ElseIf intCol = 12 Then     'COBAS
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
            End If
        Next
    Next

    spdTAT.RowHeight(-1) = 20 '30
    lblTAT.Caption = intCnt
    
    spdTAT.ReDraw = True

End Sub

'-- 06.TAT Over 현황(대기현황과 비교)
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
    intCnt = 0
        
    For intRow = 1 To spdSTB.MaxRows
        Call spdSTB.GetText(1, intRow, varTmp): strBldDtTm = varTmp
        Call spdSTB.GetText(2, intRow, varTmp): strBarNo = varTmp
        Call spdSTB.GetText(3, intRow, varTmp): strWN = varTmp
        Call spdSTB.GetText(4, intRow, varTmp): strPatInfo = varTmp
        Call spdSTB.GetText(7, intRow, varTmp): strPID = varTmp
        
        strDtTm = Format(Now, "MM/DD HH:MM")
        strOverTm = DateDiff("n", strBldDtTm, strDtTm)
        
        If strOverTm > gTatC16 Then
            spdTAT.MaxRows = spdTAT.MaxRows + 1
            Call spdTAT.SetText(2, spdTAT.MaxRows, strBarNo & "")
            Call spdTAT.SetText(3, spdTAT.MaxRows, strWN)
            Call spdTAT.SetText(4, spdTAT.MaxRows, strPatInfo)
            Call spdTAT.SetText(5, spdTAT.MaxRows, strPID)
            Call spdTAT.SetText(6, spdTAT.MaxRows, "TAT")
            Call spdTAT.SetText(7, spdTAT.MaxRows, strOverTm & " (+" & strOverTm - gTatC16 & "m)")
            intCnt = intCnt + 1
        End If
    Next

    spdTAT.RowHeight(-1) = 20 '30
    intTot = lblTAT.Caption
    lblTAT.Caption = intTot + intCnt
        
    spdTAT.ReDraw = True

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    
'    spdSTB.Width = Me.ScaleWidth * 0.25
'    spdSTB.Height = Me.ScaleHeight - 3000
'    fraSTB.Width = Me.ScaleWidth * 0.25
'
'    spdING.Left = spdSTB.Left + spdSTB.Width + 100
'    spdING.Width = Me.ScaleWidth * 0.4
'    spdING.Height = Me.ScaleHeight - 3000
'    fraING.Left = spdSTB.Left + spdSTB.Width + 100
'    fraING.Width = Me.ScaleWidth * 0.4
'
'    spdTAT.Left = spdING.Left + spdING.Width + 100
'    spdTAT.Width = Me.ScaleWidth * 0.32
'    spdTAT.Height = Me.ScaleHeight - 3000
'
'    fraTAT.Left = spdING.Left + spdING.Width + 100
'    fraTAT.Width = Me.ScaleWidth * 0.32
    
    spdSTB.Width = Me.ScaleWidth * 0.21
    spdSTB.Height = Me.ScaleHeight - 3000
    fraSTB.Width = Me.ScaleWidth * 0.21
    
    spdING.Left = spdSTB.Left + spdSTB.Width + 100
    spdING.Width = Me.ScaleWidth * 0.51
    spdING.Height = Me.ScaleHeight - 3000
    fraING.Left = spdSTB.Left + spdSTB.Width + 100
    fraING.Width = Me.ScaleWidth * 0.51
    
    spdTAT.Left = spdING.Left + spdING.Width + 100
    spdTAT.Width = Me.ScaleWidth * 0.26
    spdTAT.Height = Me.ScaleHeight - 3000
    
    fraTAT.Left = spdING.Left + spdING.Width + 100
    fraTAT.Width = Me.ScaleWidth * 0.42
    
    
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set Cn_Ser = Nothing
    
End Sub

Private Sub spdING_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp      As Variant
    Dim strBarNo    As String
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Col <> 3 Then
        Exit Sub
    End If
    
    Call spdING.GetText(Col, Row, varTmp)
    strBarNo = varTmp
    strBarNo = Mid(strBarNo, 1, Len(strBarNo) - 1)
    
    If MsgBox("검체번호 : " & strBarNo & " 를 확정처리하시겠습니까?", vbYesNo, "확정처리") = vbYes Then
              SQL = "Update " & gDB & "..TB_BCD" & vbCr
        SQL = SQL & " Set STA = 'H'" & vbCr
        SQL = SQL & " Where BCD = '" & strBarNo & "'" & vbCr
        
        Cn_Ser.Execute SQL
        
        Call cmdRefresh_Click
    
    End If
End Sub

Private Sub spdSTB_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp      As Variant
    Dim strBarNo    As String
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Col <> 2 Then
        Exit Sub
    End If
    
    Call spdSTB.GetText(Col, Row, varTmp)
    strBarNo = varTmp
    strBarNo = Mid(strBarNo, 1, Len(strBarNo) - 1)
    
    If MsgBox("검체번호 : " & strBarNo & " 를 확정처리하시겠습니까?", vbYesNo, "확정처리") = vbYes Then
              SQL = "Update " & gDB & "..TB_BCD" & vbCr
        SQL = SQL & " Set STA = 'H'" & vbCr
        SQL = SQL & " Where BCD = '" & strBarNo & "'" & vbCr
        
        Cn_Ser.Execute SQL
        
        Call cmdRefresh_Click
    
    End If
    
End Sub

Private Sub tmrMonitor_Timer()
    
    lblStatus.Caption = "TLA 모니터링중..."
    
    If lblStatus.Visible = True Then
        lblStatus.Visible = False
    Else
        lblStatus.Visible = True
    End If
    
End Sub

'-- 폼 컨트롤 초기화
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

'-- 설정파일 읽기
Private Sub GetIni()
    Dim DB_Tmp  As String * 100
    
    DB_Tmp = ""
    
    ' TLA DB 연결 ==========================================================================
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
    '======================================================================================
    

    ' 진행상태 컬럼 넓이 ==================================================================
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "WIDTH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    gWIDTH = Trim(txtTemp.Text)
    '======================================================================================
    

    ' 조회주기 ============================================================================
    DB_Tmp = ""
    Call GetPrivateProfileString("TLA", "REFRESH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    lblSec.Caption = Trim(txtTemp.Text)
    txtSec.Text = Trim(txtTemp.Text)
    lngTime = Trim(txtTemp.Text)
    '======================================================================================
    
    
    ' TAT Time 관련 =======================================================================
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
    Call GetPrivateProfileString("TLA", "C16TAT", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    txtC16TAT.Text = Trim(txtTemp.Text)
    gTatC16 = txtC16TAT.Text
    
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
    '======================================================================================
    
    
    ' 배경색 관련 =========================================================================
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
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVO", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVO = Trim(txtTemp.Text)
    If BGColor.LVO <> "" Then
        LmColor(12).CaptionBackColor = CCur(BGColor.LVO)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVOO", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVOO = Trim(txtTemp.Text)
    If BGColor.LVOO <> "" Then
        LmColor(13).CaptionBackColor = CCur(BGColor.LVOO)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVR", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVR = Trim(txtTemp.Text)
    If BGColor.LVR <> "" Then
        LmColor(14).CaptionBackColor = CCur(BGColor.LVR)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVRR", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVRR = Trim(txtTemp.Text)
    If BGColor.LVRR <> "" Then
        LmColor(15).CaptionBackColor = CCur(BGColor.LVRR)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVH = Trim(txtTemp.Text)
    If BGColor.LVH <> "" Then
        LmColor(16).CaptionBackColor = CCur(BGColor.LVH)
    End If
    
    
    '진행상태 배경색
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVO", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVO = Trim(txtTemp.Text)
    If BGColor.LVO <> "" Then
        chkO.BackColor = CCur(BGColor.LVO)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVOO", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVOO = Trim(txtTemp.Text)
    If BGColor.LVOO <> "" Then
        chkOO.BackColor = CCur(BGColor.LVOO)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVR", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVR = Trim(txtTemp.Text)
    If BGColor.LVR <> "" Then
        chkR.BackColor = CCur(BGColor.LVR)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVRR", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVRR = Trim(txtTemp.Text)
    If BGColor.LVRR <> "" Then
        chkRR.BackColor = CCur(BGColor.LVRR)
    End If
    
    DB_Tmp = ""
    Call GetPrivateProfileString("COLOR", "LVH", "", DB_Tmp, 100, App.Path & "\TLA.ini")
    txtTemp.Text = Trim(DB_Tmp)
    BGColor.LVH = Trim(txtTemp.Text)
    If BGColor.LVH <> "" Then
        chkH.BackColor = CCur(BGColor.LVH)
    End If
    '======================================================================================

End Sub

'-- 주운영 서버 접속 (MS-SQL)
Public Function Connect_PRServer() As Boolean

    Connect_PRServer = False
        
On Error GoTo errFind
    
    Set Cn_Ser = New ADODB.Connection
    
    With Cn_Ser
        .ConnectionTimeout = 25
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = gIP
        .Properties("Initial Catalog").Value = gDB
        .Properties("User ID").Value = gUID
        .Properties("Password").Value = gPWD
        .Open
    End With
    
    '연결 성공
    Connect_PRServer = True
    
    Exit Function
 
errFind:
    '연결 실패
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

