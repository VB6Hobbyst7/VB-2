VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SANSOFT"
   ClientHeight    =   10290
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInterface.frx":554A
   ScaleHeight     =   10290
   ScaleWidth      =   16080
   WindowState     =   2  '최대화
   Begin VB.Frame fraHidden 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hidden"
      Height          =   6405
      Left            =   6330
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   13845
      Begin FPSpreadADO.fpSpread spdAllergy 
         Height          =   6945
         Left            =   7230
         TabIndex        =   94
         Top             =   4410
         Width           =   7005
         _Version        =   524288
         _ExtentX        =   12356
         _ExtentY        =   12250
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   11
         MaxRows         =   66
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":588C
         UserResize      =   0
      End
      Begin VB.ListBox List1 
         Height          =   240
         Left            =   4740
         TabIndex        =   91
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Frame fraBrain 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Brain"
         Height          =   495
         Left            =   2850
         TabIndex        =   80
         Top             =   5370
         Width           =   2655
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "전체"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   83
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "대기"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   1
            Left            =   930
            TabIndex        =   82
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "완료"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   2
            Left            =   1740
            TabIndex        =   81
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Frame fraSelexOn 
         Caption         =   "SelexOn"
         Height          =   825
         Left            =   210
         TabIndex        =   76
         Top             =   5160
         Width           =   1155
         Begin VB.Timer tmrSelexOn 
            Left            =   300
            Top             =   300
         End
      End
      Begin VB.Frame fraWorkList 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   6630
         TabIndex        =   71
         Top             =   2220
         Visible         =   0   'False
         Width           =   3825
         Begin HSCotrol.CButton cmdWorkSave 
            Height          =   495
            Left            =   60
            TabIndex        =   72
            Top             =   90
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   873
            BackColor       =   4210752
            Caption         =   "현재화면 워크저장"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
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
            HoverColor      =   0
         End
         Begin HSCotrol.CButton cmdWorkLoad 
            Height          =   495
            Left            =   1830
            TabIndex        =   73
            Top             =   90
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   873
            BackColor       =   4210752
            Caption         =   "저장워크 불러오기"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
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
            HoverColor      =   0
         End
      End
      Begin VB.Frame fraBarcode 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   6630
         TabIndex        =   63
         Top             =   1470
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox txtOldBarNum 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   68
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox txtBarNum 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2190
            TabIndex        =   67
            Text            =   "123456789012345"
            Top             =   210
            Width           =   2175
         End
         Begin VB.CheckBox chkAdd 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "검체번호 변경"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   60
            TabIndex        =   66
            Top             =   60
            Width           =   1455
         End
         Begin HSCotrol.CButton cmdBarFind 
            Height          =   495
            Left            =   5550
            TabIndex        =   64
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            BackColor       =   4210752
            Caption         =   "검체찾기"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
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
            HoverColor      =   0
         End
         Begin HSCotrol.CButton cmdBarReg 
            Height          =   495
            Left            =   4440
            TabIndex        =   65
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            BackColor       =   4210752
            Caption         =   "검체등록"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
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
            HoverColor      =   0
         End
         Begin VB.Label Label10 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   ">>"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1830
            TabIndex        =   70
            Top             =   360
            Width           =   285
         End
         Begin VB.Label lblRow 
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1860
            TabIndex        =   69
            Top             =   90
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Frame fraPatInfo 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '없음
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   3540
         TabIndex        =   51
         Top             =   180
         Width           =   4635
         Begin VB.Label lblStatus 
            BackStyle       =   0  '투명
            Caption         =   "오더전송"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   3240
            TabIndex        =   57
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblPatNm 
            BackStyle       =   0  '투명
            Caption         =   "홍기롱 (M/77)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   165
            Left            =   1080
            TabIndex        =   56
            Top             =   360
            Width           =   2805
         End
         Begin VB.Label lblBarcode 
            BackStyle       =   0  '투명
            Caption         =   "123456789012345"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   165
            Left            =   1080
            TabIndex        =   55
            Top             =   150
            Width           =   2805
         End
         Begin VB.Label Label14 
            BackStyle       =   0  '투명
            Caption         =   "검사상태"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   2700
            TabIndex        =   54
            Top             =   150
            Width           =   525
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "이름 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   480
            TabIndex        =   53
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   "검체번호 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   150
            Width           =   855
         End
      End
      Begin VB.Frame fraUrometer 
         Appearance      =   0  '평면
         BackColor       =   &H00BF8B59&
         Caption         =   "Urometer"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   210
         TabIndex        =   37
         Top             =   4470
         Visible         =   0   'False
         Width           =   4965
         Begin VB.ComboBox cboUro 
            Height          =   300
            Left            =   120
            Style           =   2  '드롭다운 목록
            TabIndex        =   41
            Top             =   210
            Width           =   915
         End
         Begin VB.TextBox txtUro 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   315
            Left            =   1050
            TabIndex        =   40
            Text            =   "Positive"
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtWBC 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   315
            Left            =   3450
            TabIndex        =   39
            Top             =   210
            Width           =   585
         End
         Begin VB.TextBox txtRBC 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
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
            Height          =   315
            Left            =   4260
            TabIndex        =   38
            Top             =   210
            Width           =   585
         End
         Begin HSCotrol.CButton cmdUroSet 
            Height          =   300
            Left            =   2550
            TabIndex        =   42
            Top             =   210
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            BackColor       =   12553049
            Caption         =   "적용"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
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
            HoverColor      =   65535
         End
         Begin HSCotrol.CButton cmdPos 
            Height          =   300
            Left            =   2310
            TabIndex        =   43
            Top             =   205
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   529
            BackColor       =   33023
            Caption         =   "P"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdNeg 
            Height          =   300
            Left            =   2070
            TabIndex        =   44
            Top             =   205
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   529
            BackColor       =   16744576
            Caption         =   "N"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3240
            TabIndex        =   46
            Top             =   300
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   45
            Left            =   4080
            TabIndex        =   45
            Top             =   300
            Width           =   120
         End
      End
      Begin VB.Frame fraBIOLYTE 
         Caption         =   "BIOLYTE"
         Height          =   825
         Left            =   1470
         TabIndex        =   36
         Top             =   5190
         Width           =   1155
         Begin VB.Timer tmrBIOLYTE 
            Left            =   240
            Top             =   240
         End
      End
      Begin VB.Frame fraEPOC 
         Caption         =   "EPOC"
         Height          =   795
         Left            =   210
         TabIndex        =   34
         Top             =   2850
         Width           =   3765
         Begin VB.FileListBox FileEPOC 
            Height          =   450
            Left            =   840
            Pattern         =   "*.txt"
            TabIndex        =   35
            Top             =   180
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Timer tmrEPOC 
            Left            =   330
            Top             =   180
         End
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   2160
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame fraBIT 
         BackColor       =   &H00BF8B59&
         Caption         =   "BIT Json"
         Height          =   645
         Left            =   210
         TabIndex        =   27
         Top             =   2160
         Width           =   3945
         Begin VB.TextBox txtFrNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1770
            TabIndex        =   29
            Text            =   "0000"
            Top             =   180
            Width           =   765
         End
         Begin VB.TextBox txtToNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2820
            TabIndex        =   28
            Text            =   "0999"
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   2610
            TabIndex        =   33
            Top             =   270
            Width           =   150
         End
         Begin VB.Label lblSlipCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            Caption         =   "L20"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   600
            TabIndex        =   32
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label3 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            Caption         =   "00I"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   1260
            TabIndex        =   31
            Top             =   210
            Width           =   465
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "SLIP"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   450
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Timer tmrDBConn 
         Left            =   210
         Top             =   930
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1080
         Top             =   930
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1500
         Top             =   930
      End
      Begin VB.Timer tmrConn 
         Left            =   660
         Top             =   930
      End
      Begin VB.Frame fraVision 
         BackColor       =   &H00BF8B59&
         Caption         =   " VISION "
         Height          =   645
         Left            =   300
         TabIndex        =   23
         Top             =   3720
         Width           =   3765
         Begin VB.TextBox txtRCnt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1020
            TabIndex        =   48
            Text            =   "1"
            Top             =   210
            Width           =   525
         End
         Begin VB.CommandButton cmdGetRslt 
            Appearance      =   0  '평면
            Caption         =   "결과받기"
            Height          =   315
            Left            =   1620
            TabIndex        =   47
            Top             =   210
            Width           =   1155
         End
         Begin VB.TextBox txtLastSeq 
            Appearance      =   0  '평면
            Height          =   315
            Left            =   2880
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과갯수"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   4
            Left            =   150
            TabIndex        =   49
            Top             =   270
            Width           =   750
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtSeqNo 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   3570
         TabIndex        =   22
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtPosNo 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   2910
         TabIndex        =   21
         Top             =   1020
         Width           =   645
      End
      Begin VB.TextBox txtRackNo 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   2220
         TabIndex        =   20
         Top             =   1020
         Width           =   675
      End
      Begin VB.Timer tmrQ 
         Left            =   1200
         Top             =   210
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1650
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   150
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   720
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   2820
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":FF02
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1049C
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":10A36
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":10FD0
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":11862
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":119BC
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":11B16
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":11C70
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1254A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   495
         Left            =   10260
         TabIndex        =   58
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   " 결과출력"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":12E24
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton cmdDelete 
         Height          =   495
         Left            =   8910
         TabIndex        =   59
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   " 결과삭제"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":12F7E
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton CButton1 
         Height          =   495
         Left            =   11730
         TabIndex        =   60
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "화면정리"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":130D8
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton cmdBarcode 
         Height          =   405
         Left            =   4950
         TabIndex        =   61
         Top             =   1590
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "검체등록/찾기 ▷"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   8421504
         HoverColor      =   0
      End
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   405
         Left            =   4920
         TabIndex        =   62
         Top             =   2310
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "워크저장/로드 ▷"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   12632256
         HoverColor      =   0
      End
      Begin FPSpreadADO.fpSpreadPreview spdPrv 
         Height          =   2415
         Left            =   10890
         TabIndex        =   90
         Top             =   2340
         Width           =   2265
         _Version        =   524288
         _ExtentX        =   3995
         _ExtentY        =   4260
         _StockProps     =   96
         BorderStyle     =   1
         AllowUserZoom   =   -1  'True
         GrayAreaColor   =   8421504
         GrayAreaMarginH =   720
         GrayAreaMarginType=   0
         GrayAreaMarginV =   720
         PageBorderColor =   8388608
         PageBorderWidth =   2
         PageShadowColor =   0
         PageShadowWidth =   2
         PageViewPercentage=   100
         PageViewType    =   0
         ScrollBarH      =   1
         ScrollBarV      =   1
         ScrollIncH      =   360
         ScrollIncV      =   360
         PageMultiCntH   =   1
         PageMultiCntV   =   1
         PageGutterH     =   -1
         PageGutterV     =   -1
         ScriptEnhanced  =   0   'False
      End
      Begin FPSpreadADO.fpSpread spdAllergy_Color 
         Height          =   7365
         Left            =   7560
         TabIndex        =   93
         Top             =   4800
         Width           =   8475
         _Version        =   524288
         _ExtentX        =   14949
         _ExtentY        =   12991
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   11
         MaxRows         =   66
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":1435A
         UserResize      =   0
      End
      Begin VB.Image imgResult 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   1890
         Left            =   9510
         Top             =   2790
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "Label2"
         Height          =   225
         Left            =   4740
         TabIndex        =   92
         Top             =   3930
         Width           =   3345
      End
      Begin VB.Image imgOn 
         Height          =   480
         Left            =   690
         Picture         =   "frmInterface.frx":23736
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image imgOff 
         Height          =   480
         Left            =   150
         Picture         =   "frmInterface.frx":24000
         Top             =   1440
         Width           =   480
      End
   End
   Begin FPSpreadADO.fpSpread spdResult 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   12450
      TabIndex        =   75
      Tag             =   "20001"
      Top             =   870
      Width           =   8820
      _Version        =   524288
      _ExtentX        =   15558
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      MaxCols         =   13
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmInterface.frx":248CA
      VisibleCols     =   3
      VisibleRows     =   10
      CellNoteIndicatorColor=   16576
      HighlightAlphaBlendColor=   16755285
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   30
      Left            =   10290
      TabIndex        =   26
      Top             =   2340
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      Format          =   125501441
      CurrentDate     =   44029
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '없음
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   16080
      TabIndex        =   19
      Top             =   0
      Width           =   16080
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '아래 맞춤
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   16020
      TabIndex        =   2
      Top             =   9615
      Visible         =   0   'False
      Width           =   16080
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   12270
         TabIndex        =   77
         Top             =   -30
         Width           =   1185
         Begin VB.OptionButton optFile 
            Caption         =   "REAL"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optFile 
            Caption         =   "RCV"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   78
            Top             =   150
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdRcvView 
         Caption         =   "로그보기"
         Height          =   525
         Left            =   13500
         TabIndex        =   18
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "지우기"
         Height          =   525
         Left            =   11310
         TabIndex        =   12
         Top             =   60
         Width           =   885
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   525
         Left            =   19920
         TabIndex        =   11
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   525
         Left            =   19320
         TabIndex        =   10
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   525
         Left            =   18720
         TabIndex        =   9
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   525
         Left            =   18120
         TabIndex        =   8
         Top             =   60
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   525
         Left            =   17520
         TabIndex        =   7
         Top             =   60
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         BackColor       =   &H00C0FFFF&
         Height          =   525
         Left            =   14490
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   60
         Width           =   2205
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "보내기"
         Height          =   525
         Left            =   16740
         TabIndex        =   5
         Top             =   60
         Width           =   765
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   60
         Width           =   10245
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "받기"
         Height          =   525
         Left            =   10320
         TabIndex        =   3
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraWorkInfo 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      TabIndex        =   13
      Top             =   150
      Width           =   5445
      Begin VB.Frame fraJWINFO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   960
         TabIndex        =   86
         Top             =   -60
         Visible         =   0   'False
         Width           =   2055
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "전체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   89
            Top             =   120
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "입원"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   690
            TabIndex        =   88
            Top             =   120
            Width           =   645
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "외래"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   1350
            TabIndex        =   87
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.CheckBox chkSave 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장된 결과 포함 조회"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   3180
         TabIndex        =   84
         Top             =   60
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   60
         TabIndex        =   14
         ToolTipText     =   "조회 시작일"
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   125501441
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         ToolTipText     =   "조회 종료일"
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   125501441
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   375
         Left            =   3120
         TabIndex        =   0
         ToolTipText     =   "워크리스트를 조회한다"
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "워크조회"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   16711680
      End
      Begin HSCotrol.CButton cmdAll 
         Height          =   375
         Left            =   4020
         TabIndex        =   25
         ToolTipText     =   "조회된 워크리스트를 일괄등록한다."
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "일괄등록"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   16711680
      End
      Begin HSCotrol.CButton cmdMatch 
         Height          =   405
         Left            =   4920
         TabIndex        =   85
         Top             =   280
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   714
         BackColor       =   12553049
         Caption         =   "M"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   1470
         TabIndex        =   17
         Top             =   390
         Width           =   150
      End
      Begin VB.Label lblSch 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   " 조회기간"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   840
         WordWrap        =   -1  'True
      End
   End
   Begin FPSpreadADO.fpSpread spdWork 
      CausesValidation=   0   'False
      Height          =   8115
      Left            =   60
      TabIndex        =   74
      Tag             =   "20001"
      ToolTipText     =   "검체번호를 더블클릭하면 우측으로 이동합니다."
      Top             =   900
      Width           =   5460
      _Version        =   524288
      _ExtentX        =   9631
      _ExtentY        =   14314
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowHoriz   =   0   'False
      MaxCols         =   23
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   4
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmInterface.frx":2535C
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   1
      ShowScrollTips  =   1
      CellNoteIndicatorColor=   16576
   End
   Begin FPSpreadADO.fpSpread spdOrder 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   5550
      TabIndex        =   50
      Tag             =   "20001"
      Top             =   90
      Width           =   16020
      _Version        =   524288
      _ExtentX        =   28258
      _ExtentY        =   15690
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BackColorStyle  =   3
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16316655
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   22
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmInterface.frx":26160
      UnitType        =   0
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
      CellNoteIndicatorColor=   16777152
      HighlightAlphaBlendColor=   16751001
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sStartTime       As Date
Public sStartDate       As Date

Dim strOldBarno         As String
Dim gMnuIdx             As Integer
Dim AckOn               As Boolean
Dim Sample_Seq          As String
Dim aMod                As String
Dim iIID                As String

Private strA1cIntBase      As String   '수신한 장비기준 검사명
Private strA1cResult       As String   '수신한 결과(정성)

    
Private Sub chkAdd_Click()
    
    If chkAdd.Value = "1" Then
        lblRow.Visible = True
        txtOldBarNum.Enabled = True
        txtOldBarNum.BackColor = vbWhite
    Else
        lblRow.Visible = False
        txtOldBarNum.Enabled = False
        txtOldBarNum.BackColor = &HE0E0E0
    End If
    
End Sub

Private Sub chkSave_Click()
    
    cmdSearch.SetFocus
    
End Sub

Private Sub cmdAck_Click()
    
    txtSend.Text = txtSend.Text & ACK

End Sub

Private Sub cmdAll_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    Dim strRegWorkNo    As String
    Dim varRegWorkNo    As Variant
    
    strRegWorkNo = ""
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To spdOrder.MaxRows
                    spdOrder.Row = intORow
                    spdOrder.Col = colCHARTNO
                    If strBarno = GetText(spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    strRegWorkNo = strRegWorkNo & CStr(intWRow) & ","
                    spdOrder.MaxRows = spdOrder.MaxRows + 1
                    intRow = spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                    Next

                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To spdOrder.MaxCols
                            spdOrder.Row = 0
                            spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(spdOrder.Text) Then
                                Call SetSPDOrder(spdOrder, spdOrder.MaxRows, spdOrder.MaxRows, intOCol, intOCol)
                            End If
                        Next
                    Next

                    spdOrder.RowHeight(-1) = gROWHEIGHT
                End If
            End If
        Next
        '매번 전체 선택을 하지는 않는다.
        '.MaxRows = 0
    End With

    ' 선택한 번호를 찾아 지운다.
'    If strRegWorkNo <> "" And Len(strRegWorkNo) > 1 Then
'        strRegWorkNo = Mid(strRegWorkNo, 1, Len(strRegWorkNo) - 1)
'        varRegWorkNo = Split(strRegWorkNo, ",")
'        For intWRow = UBound(varRegWorkNo) To 0 Step -1
'            Call DeleteRow(spdWork, varRegWorkNo(intWRow), varRegWorkNo(intWRow))
'            spdWork.MaxRows = spdWork.MaxRows - 1
'        Next
'    End If
    
    '체크된 번호를 지운다
    For intWRow = spdWork.MaxRows To 1 Step -1
        spdWork.Row = intWRow
        spdWork.Col = colCHECKBOX
        If spdWork.Value = "1" Then
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
        End If
    Next
    
End Sub

Private Sub cmdBarcode_Click()

    If cmdBarcode.Caption = "검체등록/찾기 ▷" Then
        fraBarcode.Visible = True
        cmdBarcode.Caption = "검체등록/찾기 ◁"
        
        fraWorkList.LEFT = fraBarcode.LEFT + fraBarcode.WIDTH + 30
    
    Else
        fraBarcode.Visible = False
        cmdBarcode.Caption = "검체등록/찾기 ▷"
    
        fraWorkList.LEFT = fraBarcode.LEFT
    
    End If
    
    DoEvents

End Sub

Private Sub cmdBarFind_Click()
    Dim intRow As Integer
    
    If txtBarNum.Text = "" Then
        Exit Sub
    End If
    
    With spdOrder
        For intRow = 1 To .MaxRows
            If txtBarNum.Text = GetText(spdOrder, intRow, colBARCODE) Then
                Call spdActiveCell(spdOrder, intRow, colBARCODE)
                Exit For
            End If
        Next
    End With
    
End Sub

Private Sub cmdBarReg_Click()
    
    If txtBarNum.Text <> "" Then
        Call txtBarNum_KeyDown(vbKeyReturn, 0)
    End If
    
End Sub

Private Sub cmdClear_Click()

'
'Call WinExec("C:\IF_EPOC\COM\EPOC.exe 3 C:\IF_EPOC\COM\LOG", sw_normal)
'FileEPOC.PATH = "C:\IF_EPOC\COM\LOG"
'
'Exit Sub


    Call frmClear
    
End Sub

Private Sub cmdEnq_Click()
    
    txtSend.Text = txtSend.Text & ENQ
    
End Sub

Private Sub cmdEot_Click()
    
    txtSend.Text = txtSend.Text & EOT

End Sub

Private Sub cmdEtx_Click()
    
    txtSend.Text = txtSend.Text & ETX

End Sub

Private Sub cmdGetRslt_Click()
    Dim strFirstSeq     As String
    Dim strLastSeq      As String
    Dim strSendData     As String
    
    If txtLastSeq.Text = "" Then
        Exit Sub
    End If
    
    If txtRCnt.Text = "" Then
        Exit Sub
    End If
    
    If IsNumeric(txtLastSeq.Text) And IsNumeric(txtRCnt.Text) Then
        strFirstSeq = txtLastSeq.Text
        strFirstSeq = (strFirstSeq - 1) - (txtRCnt.Text - 1)
        strLastSeq = strFirstSeq + (txtRCnt.Text - 1)
        'strSendData = Text1 & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq
        strSendData = "0" & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq
        
        Call SendWSckData(strSendData)
    End If
    
End Sub

Private Sub cmdMatch_Click()
    Dim intWRow     As Integer
    Dim intWSrcRow  As Integer
    Dim intORow     As Integer
    Dim intOSrcRow  As Integer
    Dim blnSame     As Boolean
    Dim i           As Integer
    Dim intCnt      As Integer
    Dim varItems    As Variant
    Dim intItems    As Integer
    Dim intOCol     As Integer
    
    blnSame = False
    intCnt = 0
    
    For intWRow = 1 To spdWork.MaxRows
        If GetText(spdWork, intWRow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
            intWSrcRow = intWRow
        End If
    Next
    
    If intCnt = 0 Then
        Exit Sub
    End If
    
    
    If intCnt > 1 Then
        MsgBox "워크리스트에서 하나의 검체만 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    intCnt = 0
    
    For intORow = 1 To spdOrder.MaxRows
        If GetText(spdOrder, intORow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
            intOSrcRow = intORow
            blnSame = True
            If intCnt >= 2 Then
                Exit For
            End If
        End If
    Next
    
    If blnSame = False Then
        MsgBox "결과리스트에서 대상 검체를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If intCnt > 1 Then
        MsgBox "결과리스트에서 하나의 검체만 선택하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If blnSame = True Then
        For i = colER To colSTATE
            Call SetText(spdOrder, GetText(spdWork, intWSrcRow, i), intOSrcRow, i)
            If i = colBARCODE Then
                spdOrder.BackColor = vbCyan
            End If
        Next
        
        varItems = GetText(spdWork, intWSrcRow, colITEMS)
        varItems = Split(varItems, "/")
        For intItems = 0 To UBound(varItems)
            For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                spdOrder.Row = 0
                spdOrder.Col = intOCol
                If varItems(intItems) = Trim(spdOrder.Text) Then
                    Call SetText(spdOrder, "◇", intOSrcRow, intOCol)
                End If
            Next
        Next
        
        If GetText(spdWork, intOSrcRow, colSAVESEQ) <> "" Then
            '정보수정
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT "
            SQL = SQL & "   SET HOSPDATE = '" & Trim(GetText(spdOrder, intOSrcRow, colBARCODE)) & "'   " & vbCrLf
            SQL = SQL & "     , BARCODE  = '" & Trim(GetText(spdOrder, intOSrcRow, colBARCODE)) & "'   " & vbCrLf
            SQL = SQL & "     , PID      = '" & Trim(GetText(spdOrder, intOSrcRow, colPID)) & "'       " & vbCrLf
            SQL = SQL & "     , CHARTNO  = '" & Trim(GetText(spdOrder, intOSrcRow, colCHARTNO)) & "'   " & vbCrLf
            SQL = SQL & "     , SPECIMEN = '" & Trim(GetText(spdOrder, intOSrcRow, colSPECIMEN)) & "'  " & vbCrLf
            SQL = SQL & "     , DEPT     = '" & Trim(GetText(spdOrder, intOSrcRow, colDEPT)) & "'      " & vbCrLf
            SQL = SQL & "     , INOUT    = '" & Trim(GetText(spdOrder, intOSrcRow, colINOUT)) & "'     " & vbCrLf
            SQL = SQL & "     , ERYN     = '" & Trim(GetText(spdOrder, intOSrcRow, colER)) & "'        " & vbCrLf
            SQL = SQL & "     , RETESTYN = '" & Trim(GetText(spdOrder, intOSrcRow, colRT)) & "'        " & vbCrLf
            SQL = SQL & "     , PNAME    = '" & Trim(GetText(spdOrder, intOSrcRow, colPNAME)) & "'     " & vbCrLf
            SQL = SQL & "     , PSEX     = '" & Trim(GetText(spdOrder, intOSrcRow, colPSEX)) & "'      " & vbCrLf
            SQL = SQL & "     , PAGE     = '" & Trim(GetText(spdOrder, intOSrcRow, colPAGE)) & "'      " & vbCrLf
            SQL = SQL & "     , DISKNO   = '" & Trim(GetText(spdOrder, intOSrcRow, colRACKNO)) & "'    " & vbCrLf
            SQL = SQL & "     , POSNO    = '" & Trim(GetText(spdOrder, intOSrcRow, colPOSNO)) & "'     " & vbCrLf
            SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'                                   " & vbCrLf
            SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, intOSrcRow, colEXAMDATE)) & "'  " & vbCrLf
            'SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, intOSrcRow, colEXAMTIME)) & "'  " & vbCrLf
            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intOSrcRow, colSAVESEQ)) & vbCrLf
            
            If DBExec(AdoCn_Local, SQL) Then
                '-- 성공
            End If
        End If
        
        '워크리스트에서 지우기
        Call SetText(spdWork, "0", intWSrcRow, colCHECKBOX)
        DeleteRow spdWork, intWSrcRow, intWSrcRow
        spdWork.MaxRows = spdWork.MaxRows - 1
        
        '선택버튼 언체크
        Call SetText(spdOrder, "0", intOSrcRow, colCHECKBOX)
    End If
    
End Sub


Private Sub cmdNeg_Click()

    txtUro.Text = gHOSP.NEG
    txtUro.SelStart = 0
    txtUro.SelLength = Len(txtUro.Text)
    txtUro.SetFocus
    
End Sub

Private Sub cmdPos_Click()
    
    txtUro.Text = gHOSP.POS
    txtUro.SelStart = 0
    txtUro.SelLength = Len(txtUro.Text)
    txtUro.SetFocus
    
End Sub

Private Sub cmdRcv_Click()
    Dim strBuf  As String
    
    pBuffer = txtRcv.Text
'    If Trim(txtRcv.Text) = "" Then
'        Exit Sub
'    End If
'    RcvBuffer = txtRcv.Text
'    strBuf = Replace(RcvBuffer, vbLf, "")
'    strRecvData = Split(RcvBuffer, vbCr)

    Call ReceiveProcess
    'Call RcvData

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

'Private Sub cmdRcvView_Click()
'
'    frmLogView.Show
'
'End Sub

'Private Sub cmdReceive_Click()
'    Dim strInfoPath     As String
'    Dim strRsltPath     As String
'    Dim txtFilename     As String
'
'    With CommonDialog1
'        .CancelError = True
'
'        On Error GoTo ErrHandler
'        .Flags = cdlOFNHideReadOnly
'        .InitDir = gComm.RSTPATH
'        .Filter = "XML Files (info.xml)|*.xml|All Files (*.*)|*.*|"
'        .FilterIndex = 1
'        .Filename = ""
'        .ShowOpen
'        txtFilename = .Filename
'    End With
'
'    Screen.MousePointer = 11
'
'    strInfoPath = txtFilename 'gComm.RSTPATH & "\" & "info.xml"
'
'    Call DisplayNode_Info(strInfoPath)
'
'
'    If UBound(strRecvData) > 1 Then
'        Call SerialRcvData_MULTIPLATE
'    End If
'
'    Screen.MousePointer = 0
'
'    Exit Sub
'
'ErrHandler:
'  ' 사용자가 [취소] 단추를 눌렀습니다.
'Exit Sub
'
'End Sub


'Public Sub DisplayNode_Info(asPath As String)
'
'    Dim xmlDoc          As New MSXML2.DOMDocument30
'    Dim nodeBook        As IXMLDOMElement
'    Dim nodeId          As IXMLDOMAttribute
'    Dim xNode           As MSXML2.IXMLDOMNode
'    Dim namedNodeMap    As IXMLDOMNamedNodeMap
'    Dim Child_Node      As MSXML2.IXMLDOMNodeList
'
'    Dim i, j, k         As Integer
'    Dim MsgType         As String
'
'    On Error GoTo ErrXML:
'
'    Set xmlDoc = New MSXML2.DOMDocument30
'
'    xmlDoc.async = False
'    xmlDoc.Load asPath
'    'xmlDoc.Load "D:\프로젝트\VB\__JC메디컴\새움병원_MCC\IF\XML"
'
'    k = 0
'
'    If (xmlDoc.parseError.errorCode <> 0) Then
'        Dim myErr
'        Set myErr = xmlDoc.parseError
'        MsgBox ("You have error " & myErr.reason)
'    Else
'        Set Child_Node = xmlDoc.childNodes
'        For Each xNode In Child_Node
'            If xNode.nodeType = NODE_ELEMENT Then
'                Exit For
'            End If
'        Next
'
'        Erase strRecvData
'        intBufCnt = 1
'        ReDim Preserve strRecvData(7)
'        strRecvData(intBufCnt) = "H|" & xNode.childNodes.Item(0).baseName
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Runtime" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(8).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AreaUnderCurve" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(9).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Aggregation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(10).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Velocity" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(11).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AdditionalInformation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(12).childNodes.Item(1).nodeTypedValue
'        intBufCnt = intBufCnt + 1
'
'
'        Set Child_Node = Nothing
'
'    End If
'
'    Exit Sub
'
'ErrXML:
'    MsgBox "파일오류"
'
'    Exit Sub
'
'End Sub
'
'
'Public Sub DisplayNode_Result(asPath As String)
'
'    Dim xmlDoc As New MSXML2.DOMDocument30
'    Dim nodeBook As IXMLDOMElement
'    Dim nodeId As IXMLDOMAttribute
'    Dim xNode As MSXML2.IXMLDOMNode
'    Dim namedNodeMap As IXMLDOMNamedNodeMap
'    Dim Child_Node As MSXML2.IXMLDOMNodeList
''    Dim MsgType As String
''    Dim strBuffer As String
''    Dim intRow As Long
''    Dim varBuffer As Variant
''    Dim blnQc     As Boolean
'    Dim i, J, k, m As Integer
'    Dim ii, jj, kk  As Integer
'    Dim strOData    As String
'    Dim strRData    As String
'
'    On Error GoTo ErrXML:
''    On Error Resume Next
'
'    Set xmlDoc = New MSXML2.DOMDocument30
'
'    xmlDoc.async = False
'    xmlDoc.Load asPath
'    'xmlDoc.Load "D:\프로젝트\VB\광주포유병리과의원\참고\Result.xml"
'
'    If (xmlDoc.parseError.errorCode <> 0) Then
'        Dim myErr
'        Set myErr = xmlDoc.parseError
'        MsgBox ("You have error " & myErr.reason)
'    Else
'        Set Child_Node = xmlDoc.childNodes
'        For Each xNode In Child_Node
'            If xNode.nodeType = NODE_ELEMENT Then
'                'MsgType = xNode.nodeName
'                'If MsgType = "testinfo" Then
'                    Exit For
'                'End If
'            End If
'        Next
'
'
'        ii = 0
'        jj = 0
'        kk = 0
'
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        'PID : xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(8).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "2" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(9).childNodes.Item(1).nodeTypedValue
'        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "3" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(10).childNodes.Item(1).nodeTypedValue
'
'        For i = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Length - 1
'            ii = ii + 1
'            For J = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Length - 1
'                For k = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Length - 1
'                    If k = 0 Then
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & CStr(ii) & "|"
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).Attributes.Item(k).baseName 'xNode.childNodes.Item(0).childNodes.Item(0).baseName
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "|"
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).Attributes.Item(k).nodeTypedValue 'xNode.childNodes.Item(0).childNodes.Item(0).nodeTypedValue
'                    End If
'
'                    For m = 0 To xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Length - 1
'                        strRData = strRData & "" & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Item(m).baseName
'                        strRData = strRData & "" & "|"
'                        strRData = strRData & "" & xNode.childNodes.Item(0).childNodes.Item(0).childNodes.Item(i).childNodes.Item(J).childNodes.Item(k).Attributes.Item(m).nodeTypedValue
'                        strRData = strRData & "" & "|"
'                    Next
'                Next
'                'Debug.Print strRData
'                If strRData <> "" Then
'                    intBufCnt = intBufCnt + 1
'                    ReDim Preserve strRecvData(intBufCnt)
'                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & CStr(ii) & "|" & strRData
'                End If
'                strRData = ""
'
'            Next
'            'XXXCFTR -M3XXX
'        Next
'        Set Child_Node = Nothing
'
'    End If
'
'Exit Sub
'
'ErrXML:
'    MsgBox "파일오류"
'    Exit Sub
'
'End Sub


Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("선택한 결과를 전송하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
        With spdOrder
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = colCHECKBOX
                If .Value = 1 Then
                    Res = SaveTransData(lRow, spdOrder)
                    
                    If Res = -1 Then
                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdOrder, "저장실패", lRow, colSTATE
                    
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    
                    Else
                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdOrder, "저장완료", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                        
                    End If
                    spdOrder.Row = lRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
                
                .Value = "0"
                
            Next lRow
        End With
    End If
    
End Sub



Private Sub cmdRcvView_Click()
            
    If optFile(0).Value = True Then
        Call Shell("notepad.exe " & App.PATH & "\Log\" & gHOSP.MACHNM & "_" & Format(Now, "yyyy-mm-dd") & "_RCV.txt", vbNormalFocus)
    Else
        Call Shell("notepad.exe " & App.PATH & "\Log\" & gHOSP.MACHNM & "_" & Format(Now, "yyyy-mm-dd") & ".txt", vbNormalFocus)
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim strOpt  As String
    
    strOpt = ""
    
    If gDBTYPE <> "99" Then
        If gEMR = "BRAIN" Then
            If optSch(0).Value = True Then
                strOpt = 0      '전체
            ElseIf optSch(1).Value = True Then
                strOpt = 1      '대기
            ElseIf optSch(2).Value = True Then
                strOpt = 2      '완료
            End If
            
            Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork, Format(txtFrNo.Text, "0000"), Format(txtToNo.Text, "0000"), IIf(chkSave.Value = "1", True, False), strOpt)
        Else
            Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork, Format(txtFrNo.Text, "0000"), Format(txtToNo.Text, "0000"), IIf(chkSave.Value = "1", True, False))
        End If
    Else
        Dim i As Integer
    
        With spdWork
            .MaxRows = 10
            For i = 1 To 10
                Call SetText(spdWork, "1", i, colCHECKBOX)
                Call SetText(spdWork, Format(dtpTo.Value, "yyyy-mm-dd"), i, colHOSPDATE)
                Call SetText(spdWork, Format(dtpTo.Value, "mmddhhmmss") & CStr(i), i, colBARCODE)
                Call SetText(spdWork, "오세원" & CStr(i), i, colPNAME)
                'Call SetText(spdWork, "BLD/BIL/URO/KET/PRO/NIT/GLU/pH/S.G/LEU", i, colITEMS)
                Call SetText(spdWork, "96M", i, colITEMS)  'INHALANT
            Next
            .RowHeight(-1) = gROWHEIGHT
        End With
    End If
    
End Sub

Private Sub cmdSend_Click()
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub


Private Sub cmdUroSet_Click()
    Dim intRow      As Integer
    Dim intRow2     As Integer
    Dim intCol      As Integer
    Dim intCol2     As Integer
    Dim strBarno    As String
    Dim strIntBase  As String
    
    If txtUro.Text = "" Then
        Exit Sub
    Else
        strIntBase = mGetP(cboUro.Text, 2, "|")
        intCol = mGetP(cboUro.Text, 3, "|")
        With spdOrder
            For intRow = 1 To .MaxRows
                Call SetText(spdOrder, txtUro.Text, intRow, intCol)
                strBarno = GetText(spdOrder, intRow, colBARCODE)
                If strBarno <> "" Then
                    gRow = intRow
                    With mResult
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetText(spdOrder, "1", intRow, colCHECKBOX)
                    Call SetText(spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
                    Call SetText(spdOrder, mResult.RsltTime, intRow, colEXAMTIME)
                    Call SetText(spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
                                
                    Call ResultProcess(strBarno, strIntBase, txtUro.Text, txtUro.Text, "")
                    
'                    For intRow2 = 1 To spdResult.MaxRows
'                        spdResult.Row = intRow2
'                        spdResult.Col = colRTESTNM
'                        If spdResult.Text = Trim(mGetP(cboUro.Text, 1, "|")) Then
'                            Call SetText(spdResult, txtUro.Text, intRow2, colRMACHRESULT)
'                            Call SetText(spdResult, txtUro.Text, intRow2, colRLISRESULT)
'                            'Call SetLocalDB(intRow, intRow2, "")
'                            Call ResultProcess(strBarno, strIntBase, txtUro.Text, txtUro.Text)
'                            Exit For
'                        End If
'                    Next
                    'Call SetLocalDB(introw,
                    'Call ResultProcess(strBarno, strIntBase, txtUro.Text, txtUro.Text)
                    
                End If
            Next
        End With
    End If
    
End Sub

Private Sub cmdView_Click()
    
    If gWORKPOS = "M" Then
        If spdResult.Visible = False Then
            spdResult.Visible = True
            
            'spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
            
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
            spdResult.TOP = spdOrder.TOP
        Else
            spdResult.Visible = False
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - 200
        End If
    Else
        If spdResult.Visible = False Then
            spdResult.Visible = True
            
            'spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200
            
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
            spdResult.TOP = spdOrder.TOP
        Else
            spdResult.Visible = False
            spdOrder.WIDTH = Me.ScaleWidth - 200
            
        End If
    End If

End Sub




Private Sub cmdWorkList_Click()

    If cmdWorkList.Caption = "워크저장/로드 ▷" Then
        fraWorkList.Visible = True
        cmdWorkList.Caption = "워크저장/로드 ◁"
    Else
        fraWorkList.Visible = False
        cmdWorkList.Caption = "워크저장/로드 ▷"
    End If
    
    If fraBarcode.Visible = True Then
        fraWorkList.LEFT = fraBarcode.LEFT + fraBarcode.WIDTH + 30
    Else
        fraWorkList.LEFT = fraBarcode.LEFT
    End If
    
End Sub

Private Sub cmdWorkLoad_Click()
    Dim strPath  As String
    Dim TextLine
    Dim strBuffer
    Dim strCount    As String
    
    If spdOrder.MaxRows > 0 Then
        If MsgBox("현재 화면을 지우고 워크리스트를 불러오겠습니까?", vbYesNo + vbInformation, "워크리스트 불러오기") = vbNo Then
            Exit Sub
        End If
    End If
    
    
    With CommonDialog1
      .CancelError = True
      On Error GoTo ErrHandler
      .Flags = cdlOFNHideReadOnly
      .InitDir = App.PATH & "\WorkList"
      .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"
      .FilterIndex = 1
      .Filename = ""
      .ShowOpen
      strPath = .Filename
    End With
    
    Open strPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, TextLine
        strBuffer = strBuffer & TextLine & vbCr & vbLf
    Loop
    Close #1
 
    'strCount = strPath
    strCount = mGetP(mGetP(mGetP(strPath, 2, "WL_"), 3, "_"), 1, ".")
    
    spdOrder.MaxRows = strCount
    
    With spdOrder
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .Clip = strBuffer
        .ClipboardPaste
        .BlockMode = False
    End With
    
Exit Sub
ErrHandler:
                        
End Sub

Private Sub cmdWorkSave_Click()
    Dim strBuffer As String
    
    If spdOrder.MaxRows < 1 Then
        Exit Sub
    End If
    
    Call spdOrder.SetSelection(1, 1, spdOrder.MaxCols, spdOrder.MaxRows)
    '클립보드 카피
    spdOrder.ClipboardCopy
    
    strBuffer = Clipboard.GetText()
    
    Call SetWorkData(strBuffer, spdOrder.MaxRows)

End Sub

Private Sub ReceiveProcess()
    Dim vBuffers    As Variant
                    
    '>> RS232C
    If gComm.COMTYPE = "1" Then
        
        Select Case UCase(gHOSP.MACHNM)
            Case "AFIAS2":          Call Phase_Serial_AFIAS2
            
            Case "DOTTO2000":       Call Phase_Serial_DOTTO2000 'PPC125
            Case "UROMETER120":     Call Phase_Serial_UROMETER120
            Case "UROMETER720":     Call Phase_Serial_UROMETER720
            Case "SELEXON":         Call Phase_Serial_SELEXON
            Case "AU480":           Call Phase_Serial_AU480
            Case "XL1000I":         Call Phase_Serial_XL1000I
            Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
            Case "MICROS60":        Call Phase_Serial_MICROS60
            Case "LTC52":           Call Phase_Serial_LTC52
            Case "UROMETER":        Call Phase_Serial_UROMETER720
            Case "ISMART30":        Call Phase_Serial_ISMART30
            Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
            Case "ARKRAY":          Call Phase_Serial_ARKRAY
            Case "EPOC":            Call Phase_Serial_EPOC
            Case "URINSCAN":        Call Phase_Serial_URINSCAN
            Case "BIOLYTE":         Call Phase_Serial_BIOLYTE
            Case "HORIBA":          Call Phase_Serial_HORIBA
            Case "MINIVIDAS":       Call Phase_Serial_MINIVIDAS
            Case "THUNDERBOLT":     Call Phase_Serial_THUNDERBOLT
            Case "XN1000":          Call Phase_Serial_XN1000
            Case "MEDONIC":         Call Phase_Serial_MEDONIC
            Case "RP500":           Call Phase_Serial_RP500
            Case "AVL9180":         Call Phase_Serial_AVL9180
            Case "CA800_ASTM":      Call Phase_Serial_CA800_ASTM
            Case "CA800":           Call Phase_Serial_CA800
            Case "ACCESS2":         Call Phase_Serial_ACCESS2
            Case "YUMIZEN":         Call Phase_Serial_YUMIZEN           '영인과학 HORIBA YUMIZEN H500
            Case "XP300":           Call Phase_Serial_XP300
            Case "STAGO":           Call Phase_Serial_STAGO
            Case "PATHFAST":        Call Phase_Serial_PATHFAST
            Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
            'Case "KLITE":           Call Phase_Serial_KLITE
            'Case "INDIKO":          Call Phase_Serial_INDIKO
        End Select
        
    '>> SOCKET
    ElseIf gComm.COMTYPE = "2" Then
        Select Case UCase(gHOSP.MACHNM)
            Case "AFINION2":        Call Phase_TCP_AFINION2
            
            Case "BC6800":          Call TCPRcvData_BC6800
            
            Case "VISIONB":         Call Phase_TCP_VISION
            Case "BS360S":
                                    vBuffers = Split(pBuffer, vbCr)
                                    Call Sleep(200)
                                    If UBound(vBuffers) > 0 Then
                                        Call TCPRcvData_BS360S
                                    End If
                                    
            Case "BC5180":          Call TCPRcvData_BC5180
            Case "BC6200":          Call TCPRcvData_BC6200
            Case "GENEXPERT":       Call Phase_TCP_GENEXPERT
            Case "PPC300N":         Call Phase_TCP_PPC300N
            Case "KLITE":           Call Phase_TCP_KLITE
            Case "XP300":           Call Phase_TCP_XP300
            Case "YUMIZEN":         Call Phase_TCP_YUMIZEN
            
        End Select
    End If
        
End Sub

Private Sub cmdXML_Click()
    
    Dim FindFile As String

    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
    If FindFile <> "" Then
        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"
        MsgBox "XML 오더파일이 정리 되었습니다.", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Private Sub comEQP_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Select Case comEqp.CommEvent
        Case comEvReceive
            MDIIF.imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            pBuffer = comEqp.Input

            SetRawData "" & pBuffer
            
            Call ReceiveProcess

        Case comEvSend
            MDIIF.imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If

        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

    If ERMsg$ <> "" Then
        MDIIF.lblIFStatus.Caption = ERMsg$
    End If
    
End Sub


'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    Cancel = 1
'    Call cmdEnd_Click
'
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'
'    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
'
'        Close #1
'
'        If comEqp.PortOpen = True Then
'            comEqp.PortOpen = False
'        End If
'
''        Call DisConnect_Server
''
''        Call DisConnect_Local
'
''        Unload Me
'
''        End
'    End If
'
'End Sub



Private Sub GetOrder_HITACHI7180(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    Dim strSndMsg   As String
    
    intRow = -1
    GetOrder = ""

    ''Call SetCommStatus("Q", pBarNo, frmInterface.spdComStatus)
    ''Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarno, intRow)
    
        Call SetSQLData("바코드조회", strItems, "A")
        
        mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        GetOrder = ""
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = strItems
            mOrder.Order = String$(88, "0")
        
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If


        Call SendData(GetOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_HITACHI7020(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

'    'Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarno, intRow)
        
        '바코드를 사용하지 않을 경우에 사용한다.
        If gHOSP.BARUSE <> "Y" Then
            mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        End If
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_STAGO(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_STAGO(gHOSP.MACHCD, pBarno, intRow)
        
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_XN1000(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

'    Call SetCommStatus("Q", pBarNo, frmInterface.spdComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_XN1000(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_THUNDERBOLT(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

'    Call SetCommStatus("Q", pBarNo, frmInterface.spdComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_THUNDERBOLT(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 체크박스 표시
            Call SetText(spdOrder, "0", intRow, colCHECKBOX)
            
            '-- 진행상태(Order) 표시
            Call SetText(spdOrder, "오더없음", intRow, colSTATE)
        
            '-- 오더 아이템 저장
            Call SetText(spdOrder, "", intRow, colSPECIMEN)
        
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 체크박스 표시
            Call SetText(spdOrder, "1", intRow, colCHECKBOX)
            
            '-- 진행상태(Order) 표시
            Call SetText(spdOrder, "오더준비", intRow, colSTATE)
            
            '-- 오더 아이템 저장
            Call SetText(spdOrder, strItems, intRow, colSPECIMEN)
            
        End If


        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_CA800(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String
    Dim SendBuf     As String
    
    intRow = -1
    GetOrder = ""

'    Call SetCommStatus("Q", pBarNo, frmInterface.spdComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_CA800(gHOSP.MACHCD, pBarno, intRow)
        
        SendBuf = "S"
        SendBuf = SendBuf & "2"
        SendBuf = SendBuf & "21"
        SendBuf = SendBuf & "01"
        SendBuf = SendBuf & "01"
        SendBuf = SendBuf & "U"
        SendBuf = SendBuf & Format$(Date, "YYMMDD")
        SendBuf = SendBuf & Format$(Now, "HHMM")
        SendBuf = SendBuf & mOrder.RackNo
        SendBuf = SendBuf & mOrder.TubePos
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            SendBuf = SendBuf & Space(15)
            SendBuf = SendBuf & "C"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            SendBuf = SendBuf & Right(Space(15) & mOrder.BarNo, 15)
            SendBuf = SendBuf & "B"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If

        SendBuf = STX & SendBuf & ETX
        
        Call Sleep(500)
        
        Call SendData(SendBuf)

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_ACCESS2(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, pBarno, intRow)
        
        mOrder.Order = strItems
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_DOTTO2000(ByVal pBarno As String, ByVal pType As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        'Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_DOTTO2000(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더준비", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_YUMIZEN(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    'Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_YUMIZEN(gHOSP.MACHCD, pBarno, intRow)
        mOrder.Order = strItems
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            'mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            'mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String

    intRow = -1

    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        strItems = ""
        mOrder.Order = ""
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarno, intRow)

        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
            
        Else
            mOrder.NoOrder = False
        
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & strItems & ETX
        End If

        Call SendData(strOrder)
        
        '-- 현재 Row
        gRow = intRow

    End With

End Sub


'-- 장비로 전송 및 기록
Private Sub SendData(ByVal pSendData As Variant)

    '-- 전송
    comEqp.Output = pSendData
    
    MDIIF.imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    DoEvents
    
    '-- 로그기록
    Call SetRawData("[Tx]" & pSendData)

End Sub

Private Sub SendWSckData(ByVal pSendData As Variant)

    '-- 전송
    wSck.SendData pSendData
    
'    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
'    If tmrSend.Enabled = False Then
'        tmrSend.Enabled = True
'    Else
'        tmrSend.Enabled = False
'        tmrSend.Enabled = True
'    End If
'    DoEvents
    
    '-- 로그기록
    Call SetRawData("[Tx]" & pSendData)

End Sub


Private Sub TCPRcvData_KLITE()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20190611090403||ORU^R01|TR14-009|P|2.4||||||ASCII
                    strHeader = mGetP(strRcvBuf, 10, "|")
                    strHeaderType = mGetP(strRcvBuf, 18, "|")
                    
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||" & strHeaderType & "|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|" & strHeader & "|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    'If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    'End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- 참고치
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- 남자참고치를 기본으로 한다
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- 결과Row 추가
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- 소수점 처리
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- 결과판정
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- 진행상태 표시("결과")
                            SetText .spdOrder, "장비결과", gRow, colSTATE
    
                            '-- 메인화면 결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
                            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '순번
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '장비채널
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS결과
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
                            
                            '-- 이전결과 조회
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '이전결과
                            
                            '-- H/L 색깔표시
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- 로컬 저장
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = gROWHEIGHT
        
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT

            End Select
        Next
    
        '## DB에 결과저장
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'Public Sub TCPRcvData_BS200()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim varResult       As Variant
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'    Dim intCnt          As Integer
'
'    Dim strOrderCode    As String   '처방코드
'    Dim strTestCode     As String   '검사코드
'    Dim strTestSubCode  As String   '검사코드
'    Dim strTestName     As String   '검사명
'    Dim strSeqNo        As String   '로컬DB 검사Seq
'
'    Dim strTmp          As String
'
'    Dim strTGResult     As String
'    Dim strCHOLResult   As String
'    Dim strHDLResult    As String
'    Dim intCol          As Integer
'
'    Dim blnResult       As Boolean
'
'    Dim strRstRow       As String   '결과스프레드 현재 Row
'    Dim strDecYN        As String   '결과판정여부
'    Dim strJudge        As String   '결과판정
'
'    Dim strQCData       As String
'    Dim i               As Integer
'    Dim Res             As Integer
'    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'
'    Dim strSndBuffer    As String
'
'    'eGFR
'    Dim strCREA     As String
'    Dim strGFR      As String
'    Dim strSex      As String
'    Dim strAge      As String
'
'    Dim strHbA1c    As String
'    Dim strIFCC     As String
'    Dim streAG      As String
'    Dim strTotA1C   As String
'
'
'    blnResult = False
'
'    '-- LDL 계산용
''    strTGResult = ""
''    strCHOLResult = ""
''    strHDLResult = ""
'
'    strCREA = ""
'    strGFR = ""
'
'    strHbA1c = ""
'    strIFCC = ""
'    'strADAG = ""
'    streAG = ""
'    strTotA1C = ""
'
'    With frmInterface
'        For intCnt = 0 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'            'SetRawData "[Rcv]" & strRcvBuf
'
'            strType = mGetP(strRcvBuf, 1, "|")
'
'            Select Case strType
'                Case "MSH"
'                    'Corp.name(3)           : MINDRAY
'                    'Device Model(4)        : BS-380
'                    'System date/time(7)    : 20130504083053
'                    'Message Type(9)        : QRY^Q02
'                    'Message ID(10)         : 1
'                    'Product(11)            : P
'                    'HL7 Version(12)        : 2.3.1
'                    'Resut Type(16)         : '' (오더), 0 (Sample) , 1 (Calib. Result)
'                    'Character Encoding(18) : ASCII
'
'                    mOrder.BSMaker = mGetP(strRcvBuf, 3, "|")
'                    mOrder.BSMchNm = mGetP(strRcvBuf, 4, "|")
'                    mOrder.BSMType = mGetP(strRcvBuf, 9, "|")
'                    mOrder.BSDtTm = Format(Now, "yyyymmddhhmmss")
'
'                    With mOrder
'                        .MSHCorpName = mGetP(strRcvBuf, 3, "|")
'                        .MSHDeviceModel = mGetP(strRcvBuf, 4, "|")
'                        .MSHSysDateTime = mGetP(strRcvBuf, 7, "|")
'                        .MSHMessageType = mGetP(strRcvBuf, 9, "|")
'                        .MSHMessageID = mGetP(strRcvBuf, 10, "|")
'                        .MSHProduct = mGetP(strRcvBuf, 11, "|")
'                        .MSHHL7Version = mGetP(strRcvBuf, 12, "|")
'                        .MSHResultType = mGetP(strRcvBuf, 16, "|")
'                        .MSHChrEncoding = mGetP(strRcvBuf, 18, "|")
'                    End With
'
'                    Select Case mOrder.MSHMessageType
'                        '-- 검사결과수신 ACK
'                        Case "ORU^R01"  '==> ACK^R01
'                                           strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
'                            strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
'                            strSndBuffer = strSndBuffer & EB & vbCr
'
'                            SetRawData "[Tx]" & strSndBuffer
'                            wSck.SendData strSndBuffer
'                        '-- 오더요청수신
'                        Case "QRY^Q02"  '==> QCK^Q02
'
'                            strSndBuffer = ""
'
'                            With spdOrder
'                                For i = 1 To .MaxRows
'                                    If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
'                                        '-- 오더있음
'                                                       strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||QCK^Q02|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
'                                        strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
'                                        strSndBuffer = strSndBuffer & "ERR|0|" & vbCr & EB & vbCr
'                                        strSndBuffer = strSndBuffer & "QAK|SR|OK|" & vbCr
'                                        strSndBuffer = strSndBuffer & EB & vbCr
'
'                                        'If wSck.State <> sckClosed Then
'                                            SetRawData "[Tx]" & strSndBuffer
'                                            wSck.SendData strSndBuffer
'                                        'End If
'                                        Exit For
'                                    End If
'                                Next
'                            End With
'
'                            '-- 오더없음
'                            If strSndBuffer = "" Then
'                                               strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||QCK^Q02|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
'                                strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
'                                strSndBuffer = strSndBuffer & "ERR|0|" & vbCr & EB & vbCr
'                                strSndBuffer = strSndBuffer & "QAK|SR|NF|" & vbCr
'                                strSndBuffer = strSndBuffer & EB & vbCr
'
'                                'If wSck.State <> sckClosed Then
'                                    SetRawData "[Tx]" & strSndBuffer
'                                    wSck.SendData strSndBuffer
'                                'End If
'                            End If
'
'                        '-- 오더 전송
'                        Case "ACK^Q03"
'                            '-- 최초이후전송
'                            Call GetOrder_BS200(strBarno, gHOSP.RSTTYPE)
'
'                    End Select
'
'                Case "QRD"
'                    'QRD|20180611153634|R|D|1|||RD|0019|OTH|||T|
'                    'Qry Time(2)                    : 20180611153634
'                    'Qry Format Code(3)             : R
'                    'Qry Priority(4)                : D
'                    'Quantity Limited Request(8)    : RD
'                    'Sample Barcode(9)              : 0019
'                    'What Subject Filter(10)        : OTH
'                    'Query Results Level(13)        : T
'
'                    'QRD|20190828133858|R|D|1|||RD||OTH|||T|
'
'                    With mOrder
'                        .QRDQryTime = mGetP(strRcvBuf, 2, "|")
'                        .QRDQryFormatCode = mGetP(strRcvBuf, 3, "|")
'                        .QRDQryPriority = mGetP(strRcvBuf, 4, "|")
'                        .QRDNum = mGetP(strRcvBuf, 5, "|")
'                        .QRDQLRequest = mGetP(strRcvBuf, 8, "|")
'                        .QRDSampleBarcode = mGetP(strRcvBuf, 9, "|")
'                        .QRDWSFilter = mGetP(strRcvBuf, 10, "|")
'                        .QRDQryResultLevel = mGetP(strRcvBuf, 13, "|")
'                    End With
'
'                Case "QRF"
'                    'QRF|BS-380|19000101000000|20130504083053|||RCT|COR|ALL||
'                    'Which Date/Time Qualifier          : RCT
'                    'Which Date/Time Status Qualifier   : COR
'                    'Date/Time Selection Qualifier      : ALL
'
'                    mOrder.BSModel = mGetP(strRcvBuf, 2, "|")
'                    mOrder.BSSTime = mGetP(strRcvBuf, 3, "|")
'                    mOrder.BSETime = mGetP(strRcvBuf, 4, "|")
'                    mOrder.BSQRF = strRcvBuf
'                    mOrder.Seq = mGetP(strRcvBuf, 5, "|")
'
'                    With mOrder
'                        .QRFProduct = mGetP(strRcvBuf, 2, "|")
'                        .QRFWherStartDtTm = mGetP(strRcvBuf, 3, "|")
'                        .QRFWherEndDtTm = mGetP(strRcvBuf, 4, "|")
'                        .QRFWhichDtTmQualifier = mGetP(strRcvBuf, 7, "|")
'                        .QRFWhichStatusQualifier = mGetP(strRcvBuf, 8, "|")
'                        .QRFDtTmSelecQualifier = mGetP(strRcvBuf, 9, "|")
'                    End With
'
'                    '-- 최초오더전송
'                    intSndPhase = 1
'
'                    Call GetOrder_BS200(strBarno, gHOSP.RSTTYPE)
'
'                Case "PID"
'                    mOrder.BSMType = mGetP(strRcvBuf, 2, "|")
'                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
'                    'mOrder.PName = mGetP(strRcvBuf, 5, "|")
'                    mResult.BarNo = strBarno
'                    If Trim(strBarno) <> Trim(strOldBarno) Then
'                        strOldBarno = strBarno
'
'                        With mResult
'                            .BarNo = strBarno
'                            .RsltDate = Format(Now, "yyyy-mm-dd")
'                            .RsltTime = Format(Now, "hh:mm:ss")
'                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'                        End With
'
'                    End If
'
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                Case "OBR"
'                    'OBR|28|6|CHOL|^|Serum|20180529164220|20180529164044|20180529175810|||1|1|Normal|26411|20190131||M|255.000000|18.000000|249.413219|mg/dL|||||||||||||||||||||||||||
'
'                    'Sample결과
'                    If mOrder.MSHResultType = "0" Then
'                        strSeq = Trim$(mGetP(strRcvBuf, 4, "|"))
'
'                        If strBarno = "" Then
'                            strBarno = strSeq
'                        End If
'                    Else
'                        'cal 결과는 처리안함
'                        Exit Sub
'                    End If
'
'
'                Case "OBX"
'                    strIntBase = Trim(mGetP(strRcvBuf, 4, "|"))
'                    If strIntBase = "" Then
'                        strIntBase = Trim(mGetP(strRcvBuf, 5, "|"))
'                    End If
'
'                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
'                    'strResult = Format(strResult, "0.00")
'
'                    '-- CREA 결과저장
'                    If Trim(strIntBase) = "CRE" Then
'                        strGFR = ""
'                        strResult = Format(strResult, "##0.00")
'                        strCREA = strResult
'
'                        If CCur(strResult) > 0 Then
'                            '18세 이상만 적용
'                            If IsNumeric(strCREA) And mPatient.AGE > 18 Then
'                                If mPatient.SEX = "M" Then
'                                    strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203)
'                                ElseIf mPatient.SEX = "F" Then
'                                    strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
'                                End If
'
'                                If strGFR <> "" Then
'                                    strGFR = Format(strGFR, "##0.00")
'                                    If strGFR <= 120 Then
'                                        strGFR = Round(strGFR, 2)
'                                    ElseIf strGFR > 120 Then
'                                        strGFR = "> 120"
'                                    End If
'                                End If
'                            End If
'                        Else
'                            strGFR = "Error"
'                        End If
'                    End If
'
'
''                    If Trim(strIntBase) = "A1C" Then
''                        strA1C = strResult
''                    End If
'                    If Trim(strIntBase) = "HbA1c%" Then
'                        strResult = Format(strResult, "##0.00")
'                        strHbA1c = strResult
'                    End If
'                    If Trim(strIntBase) = "IFCC" Then
'                        strResult = Format(strResult, "##0.00")
'                        strIFCC = strResult
'                    End If
''                    If Trim(strIntBase) = "ADAG" Then
''                        strADAG = strResult
''                    End If
'                    If Trim(strIntBase) = "eAG" Then
'                        strResult = Format(strResult, "##0.00")
'                        streAG = strResult
'                    End If
'
'RST:
'                    If strIntBase <> "" And strResult <> "" Then
'                        SQL = ""
'                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH  " & vbCr
'                        SQL = SQL & ", QCTemp AS DECYN                              " & vbCr
'                        SQL = SQL & "  FROM EQPMASTER                               " & vbCr
'                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'        " & vbCr
'                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'      " & vbCr
'                        '처방이 있을경우
'                        If gPatOrdCd <> "" Then
'                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ")     " & vbCr
'                            strState = "R"
'                        Else
'                            strState = ""
'                        End If
'
'
'                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                            strTestCode = Trim(RS_L.Fields("TESTCODE"))
'                            strTestName = Trim(RS_L.Fields("TESTNAME"))
'                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
'                            strQCTemp = Trim(RS_L.Fields("DECYN") & "")
'
'                            '-- 결과Row 추가
'                            strRstRow = .spdResult.DataRowCnt + 1
'                            If .spdResult.MaxRows < strRstRow Then
'                                .spdResult.MaxRows = strRstRow
'                            End If
'
'                            '소수점 처리, 결과 형태 처리
'                            strMachResult = strResult
'                            If strQCTemp = "1" Then
'                                strResult = SetResult(strResult, strIntBase)
'                            End If
'                            strJudge = SetJudge(strResult, strIntBase)
'
'                            '진행상태 표시("결과")
'                            SetText .spdOrder, "결과", gRow, colSTATE
'
'                            '결과값 표시
'                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
'                                    SetText .spdOrder, strResult, gRow, intCol
'
'                                    '서브코드
'                                    strTestSubCode = gArrEQP(intCol - colSTATE, 17)
'
'                                    Exit For
'                                End If
'                            Next
'
'                            '-- 결과 List
'                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                            SetText .spdResult, strTestSubCode, strRstRow, colRSUBCD          '검사SUB코드
'                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                            '-- 로컬 저장
'                            SetLocalDB gRow, strRstRow, "1", ""
'
'                            'strState = "R"
'
'                            '-- 결과Count
'                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                SetText .spdOrder, "1", gRow, colRCNT
'                            Else
'                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                            End If
'                        Else
'                            strState = ""
'                        End If
'                    End If
'
'                    .spdResult.RowHeight(-1) =gROWHEIGHT

'                    '-- GFR 저장
'                    If strGFR <> "" Then
'                        strIntBase = "eGFR"
'                        strResult = strGFR
'                        strGFR = ""
'                        GoTo RST
'                    End If
'
'                    If strHbA1c <> "" And strIFCC <> "" And streAG <> "" Then
'                        '소견저장일때
'                        strTotA1C = ""
'                        'strTotA1C = strTotA1C & "A1C : " & strA1C & vbCrLf
'                        strTotA1C = strTotA1C & "HbA1c% : " & strHbA1c & vbCrLf
'                        strTotA1C = strTotA1C & "IFCC : " & strIFCC & vbCrLf
'                        'strTotA1C = strTotA1C & "ADAG : " & strADAG & vbCrLf
'                        strTotA1C = strTotA1C & "eAG : " & streAG & vbCrLf
'                        strTotA1C = Mid(strTotA1C, 1, 254)
'
'                        'HbA1C결과만 저장한다.
'                        strTotA1C = strHbA1c
'
'                        'strA1C = ""
'                        strHbA1c = ""
'                        strIFCC = ""
'                        'strADAG = ""
'                        streAG = ""
'
'                        strIntBase = "A1C"  'C3825
'                        strResult = strTotA1C
'
'                        GoTo RST
'                    End If
'            End Select
'        Next
'    'OBX|1|NM|HB|Hemoglobin|175.110862|μmol/L|-|N|||F||175.110862|20190829101649|||0||
'    'OBX|2|NM|A1C|Hemoglobin A1c|6.920507|μmol/L|-|N|||F||6.920507|20190829101649|||0||
'    'OBX|3|NM||HbA1c%|5.755654||-|N|||F||5.755654||||||
'    'OBX|4|NM||IFCC|39.409297||-|N|||F||39.409297||||||
'    'OBX|5|NM||ADAG|6.561490||-|N|||F||6.561490||||||
'    'OBX|6|NM||eAG|118.487267|mg/dL|-|N|||F||118.487267||||||
'
'
'
'        '## DB에 결과저장
'        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'            Res = SaveTransData(gRow, spdOrder)
'
'            If Res = -1 Then
'                '-- 저장 실패
'                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                SetText .spdOrder, "저장실패", gRow, colSTATE
'            Else
'                '-- 저장 성공
'                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                SetText .spdOrder, "저장완료", gRow, colSTATE
'                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                      SQL = "Update PATRESULT Set " & vbCrLf
'                SQL = SQL & " sendflag = '2' " & vbCrLf
'                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                If DBExec(AdoCn_Local, SQL) Then
'                    '-- 성공
'                End If
'            End If
'            strState = ""
'
'        End If
'    End With
'
'End Sub

Private Sub TCPRcvData_GENEXPERT()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    Dim strSend         As String
    
    Dim strMTB          As String
    Dim strRIF          As String
    Dim strCDIF         As String
    Dim str027          As String
    Dim strCarbaRPos    As String
    Dim strCarbaRNeg    As String
    
    Dim strMTBRIFCMT    As String
    Dim strCarbaRCMT    As String
    'Dim strCarbaRNeg    As String
    Dim strMachNum      As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = Mid(strRcvBuf, 1, 1)
            If IsNumeric(strType) Then
                strType = Mid(strRcvBuf, 2, 1)
            End If

            Select Case strType
                Case "H"
                    mResult.CARBAR_CMTCD = ""
                    mResult.MTBRIF_CMTCD = ""
                    mResult.CMNTCD = ""
                Case "P"
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                    
                    If strBarno = "" Then
                        Exit Sub
                    End If
                
'''                    If Trim(strBarno) <> Trim(strOldBarno) Then
'''                        '-- 결과정보
'''                        With mResult
'''                            .BarNo = strBarno
'''                            .RsltDate = Format(Now, "yyyy-mm-dd")
'''                            .RsltTime = Format(Now, "hh:mm:ss")
'''                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'''                        End With
'''                    End If
'''
'''                    strOldBarno = strBarno
'''
'''                    '-- 결과환자정보
'''                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''
'''                    If gRow <= 0 Then
'''                        Exit Sub
'''                    End If
                
                
                Case "R"
                    'R|1|^MTB-RIF^^MTB^Xpert MTB-RIF Assay G4^6^MTB^|DETECTED HIGH^|||||F||<None>|20190819150251|20190819164413|Cepheid-642628D^820753^723731^785423426^24912^20210321
'A1: 699607
'A2: 699606
'A3: 699605
'A4: 699604
'B4: 723731
'B3: 723731
'B2: 723715
'B1: 722171
'
                    
                    strMachNum = mGetP(mGetP(strRcvBuf, 14, "|"), 3, "^")
                    mResult.EqpCd = "E13"
                    Select Case strMachNum
                        Case "699607", "699607", "699607", "699607"
                            mResult.EqpCd = "E13"
                        Case "723731", "723731", "723715", "722171"
                            mResult.EqpCd = "E14"
                    End Select
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                        
                    strOldBarno = strBarno
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strIntBase = mGetP(strRcvBuf, 3, "|")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = "" 'mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    
                    Call SetSQLData("RCV", strIntBase & ":" & strResult, "A")
                    
                    '-- MTB Ct값 찾기
'''                    If strIntBase = "^MTB-RIF^^MTB^^^Probe E^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 38 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
'''
'''                    '-- TOX Ct값 찾기
'''                    If strIntBase = "^G3^^Toxi^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 5 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
'''
'''                    '-- Carba-R 값 찾기
'''                    If strIntBase = "^Carba-R^^IMP1^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "판정불가"
'''                        End If
'''                    End If
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        'MTB
                        If strIntBase = "^MTB-RIF^^MTB^Xpert MTB-RIF Assay G4^6^MTB^" Then
                            strMTB = strResult
                        End If
                        'RIF
                        If strIntBase = "^MTB-RIF^^RIF^Xpert MTB-RIF Assay G4^6^Rif Resistance^" Then
                            strRIF = strResult
                        End If
                        
                        'Carba-R
                        'IMP
                        If strIntBase = "^Carba-R^^IMP1^Xpert Carba-R^2^IMP1^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "IMP1" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "IMP1" & "/"
                            End If
                        End If
                        'VIM
                        If strIntBase = "^Carba-R^^VIM^Xpert Carba-R^2^VIM^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "VIM" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "VIM" & "/"
                            End If
                        End If
                        'NDM
                        If strIntBase = "^Carba-R^^NDM^Xpert Carba-R^2^NDM^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "NDM" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "NDM" & "/"
                            End If
                        End If
                        'KPC
                        If strIntBase = "^Carba-R^^KPC^Xpert Carba-R^2^KPC^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "KPC" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "KPC" & "/"
                            End If
                        End If
                        'OXA48
                        If strIntBase = "^Carba-R^^OXA48^Xpert Carba-R^2^OXA48^" Then
                            If strResult = "DETECTED" Or strResult = "POS" Then
                                strCarbaRPos = strCarbaRPos & "OXA48" & "/"
                            Else
                                strCarbaRNeg = strCarbaRNeg & "OXA48" & "/"
                            End If
                        End If
                        
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                        
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT
                
                Case "L"
                    If strMTB = "NOT DETECTED" And strRIF = "" Then
                        strIntBase = "^MTB-RIF^^RIF^Xpert MTB-RIF Assay G4^6^Rif Resistance^"
                        strResult = "*"
                        
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    If strMTB = "NOT DETECTED" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되지 않았으나 결핵이 의심되면 타검사 결과를 확인하시기 바랍니다."
                    
                        mResult.MTBRIF_CMTCD = "TB2"
                    
                    ElseIf strMTB = "DETECTED VERY LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB1"
                        
                    ElseIf strMTB = "DETECTED LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB3"
                    
                    ElseIf strMTB = "DETECTED MEDIUM" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine
                    
                        mResult.MTBRIF_CMTCD = "TB4"
                    
                    ElseIf strMTB = "DETECTED HIGH" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균이 검출되어 감염병 병원체 신고대상입니다." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "결핵균 검출 시 결핵균 농도가 반정량적으로 보고됩니다." & vbNewLine
                        
                        mResult.MTBRIF_CMTCD = "TB5"
                    
                    End If
            
                    If strRIF = "DETECTED" Then
                        If strMTB = "DETECTED VERY LOW" Then
                            mResult.MTBRIF_CMTCD = "RIF1"
                            
                        ElseIf strMTB = "DETECTED LOW" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF2"
                        
                        ElseIf strMTB = "DETECTED MEDIUM" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF3"
                        
                        ElseIf strMTB = "DETECTED HIGH" Then
                            
                            mResult.MTBRIF_CMTCD = "RIF4"
                        
                        End If
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "Rifamin 내성연관 돌연변이가 검출되어 내성으로 판단됩니다." & vbNewLine
                    
                    End If
                    
                    mResult.MTBRIF_CMT = strMTBRIFCMT
                    
                    strMTB = ""
                    strRIF = ""
                    strMTBRIFCMT = ""
                    
                    If strCarbaRPos <> "" Then
                        strCarbaRPos = Mid(strCarbaRPos, 1, Len(strCarbaRPos) - 1)
                        strCarbaRPos = Replace(strCarbaRPos, "/", " ")
                        
                        strCarbaRCMT = ""
                        strCarbaRCMT = strCarbaRCMT & "[Comment]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "검출된 Carbapenemase 유전자형 : strCarbaRPos" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "환자의 검체에서 Carbapenemase 유전자가 검출되었습니다." & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "Carbapenemase-producing Enter obacteriaceae (CPE) 보균자로 판단됩니다." & vbNewLine
                        
                    Else
                        strCarbaRCMT = ""
                        strCarbaRCMT = strCarbaRCMT & "[Comment]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "검출된 Carbapenemase 유전자형 : 없음" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "본 검사는 KPC, NDM, VIM 및 OXA-48 이외의 검사에서 carbapenemase에 의해서 발생한 CRE나," & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "필요 시 CRE 선별배양검사(검사코드 : 40920)를 의뢰하시기 바랍니다." & vbNewLine
                        
                    End If
                    
                    mResult.CARBAR_CMT = strCarbaRCMT
                    strCarbaRNeg = ""
                    strCarbaRPos = ""
                    strCarbaRCMT = ""
                     
                    If mResult.MTBRIF_CMTCD <> "" Then
                        mResult.CMNTCD = mResult.MTBRIF_CMTCD
                    End If
            End Select
        Next
        
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_GENEXPERT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_PPC300N()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20190611090403||ORU^R01|TR14-009|P|2.4||||||ASCII
                    strHeader = mGetP(strRcvBuf, 10, "|")
                    strHeaderType = mGetP(strRcvBuf, 18, "|")
                    
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
'[Rx]MSH|^~\&|PKL|PKL PPC 300N|||20190807112436||ORU^R01|201908070001|p|2.3.1||||0||ASCII|||
'PID|1||||||||||||||||||||||||||||||
'OBR|1||201908070001|PKL^PKL PPC 300N||||||||||||||||||||||||||||||||||||||||||
'OBX|1|NM|1|CHOL|222|mg/dL|130.0-250.0|N|||F||0.232932|||Admin||
'
                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||" & strHeaderType & "|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|" & strHeader & "|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    'If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    'End If
                
                    strSeq = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strSeq) <> Trim(strOldBarno) Then
                        strOldBarno = strSeq
                        '-- 결과정보
                        With mResult
                            '.BarNo = strBarno
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    'strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strIntBase = mGetP(strRcvBuf, 5, "|")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    strIntResult = strResult
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If


                    .spdResult.RowHeight(-1) = gROWHEIGHT

            End Select
        Next
    
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_PPC300N" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Public Sub TCPRcvData_BC6200()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim varResult       As Variant
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim intCnt          As Integer
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strTmp          As String
    
    Dim strTGResult     As String
    Dim strCHOLResult   As String
    Dim strHDLResult    As String
    Dim intCol          As Integer
    
    Dim blnResult       As Boolean
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strSndBuffer    As String
    
On Error GoTo ErrHandle
    
    strRecvData = Split(pBuffer, vbCr)
    
    With frmInterface
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    '오더요청
                    'MSH|^~\&||Mindray|||20081120174836||ORM^O01|4|P|2.3.1||||||UNICODE
                    
                    '검사결과(Sample)
                    'MSH|^~\&||Mindray|||20111124091140||ORU^R01|1|P|2.3.1||||||UNICODE
                    
                    '검사결과(QC)
                    'MSH|^~\&||Mindray|||20111124091422||ORU^R01|1|Q|2.3.1||||||UNICODE
                    
                    strOldBarno = ""
                    mOrder.MsgType = mGetP(strRcvBuf, 9, "|")
                    mOrder.MsgCtrlID = mGetP(strRcvBuf, 10, "|")
                    mOrder.TstType = mGetP(strRcvBuf, 11, "|")
                
                    If mOrder.MsgType = "ORU^R01" Then
                        '검사결과(Sample) Response Message
                        'MSH|^~\&|LIS||||20111124091140||ACK^R01|1|P|2.3.1||||||UNICODE
                        'MSA|AA|1
                        
                        '검사결과(QC) Response Message
                        'MSH|^~\&|LIS||||20111124091422||ACK^R01|1|Q|2.3.1||||||UNICODE
                        'MSA|AA|1
                                       strSndBuffer = SB
                        strSndBuffer = strSndBuffer & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|1|" & mOrder.TstType & "|2.3.1||||||UNICODE" & vbCr
                        strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MsgCtrlID & vbCr
                        strSndBuffer = strSndBuffer & EB & vbCr
                            
                        Call SendWSckData(strSndBuffer)
                    End If
                    
                Case "ORC"
                    'ORC|RF||SampleID1||IP
                    'ORC|RF||2012090228|BL
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mOrder.BarNo = strBarno
                    
                    If mOrder.MsgType = "ORM^O01" Then
                        Call GetOrder_BC6200(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                Case "PID"
                    'PID|1||^^^^MR

                Case "PV1"
                    'PV1|1
                    
                Case "OBR"
                    'OBR|1||ste5|00001^Automated Count^99MRC|||20111101170410|||||||||||||||||HM
                    'QC
                    'OBR|1||1|00003^LJ QCR^99MRC|||20201210085221|||||||||||||||||HM||||||||User
                    'Sample
                    'OBR|1||2012100197|00001^Automated Count^99MRC|||20201210113216|||||||||||||||||HM||||||||User

                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mResult.BarNo = strBarno
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = Trim(strBarno)
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                Case "OBX"
                    'OBX|5|NM|6690-2^WBC^LN||6.58|10*9/L|4.00-10.00|N|||F
                    'QC
                    'OBX|5|NM|6690-2^WBC^LN||3.61|10*3/uL|3.10-4.10|N|||F
                    'Sample
                    'OBX|9|NM|6690-2^WBC^LN||4.57|10*3/uL|4.00-10.00|N|||F
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    If strIntBase = "" Then
                        strIntBase = Trim(mGetP(strRcvBuf, 5, "|"))
                    End If
                    
                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
                    strIntResult = strResult
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT
            End Select
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With
    
    
    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_BS240E" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

Public Sub TCPRcvData_BC6800()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    Dim Res             As Integer
    Dim intCnt          As Integer
    
    Dim strSndBuffer    As String

On Error GoTo ErrHandle
'On Error Resume Next

    strRecvData = Split(pBuffer, vbCr)
    pBuffer = ""
    With frmInterface
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    strOldBarno = ""
                    mOrder.MsgType = mGetP(strRcvBuf, 9, "|")
                    mOrder.MsgCtrlID = mGetP(strRcvBuf, 10, "|")
                    mOrder.TstType = mGetP(strRcvBuf, 11, "|")
                
                    If mOrder.MsgType = "ORU^R01" Then
                        strSndBuffer = ""
                        strSndBuffer = strSndBuffer & SB
                        strSndBuffer = strSndBuffer & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|1|" & mOrder.TstType & "|2.3.1||||||UNICODE" & vbCr
                        strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MsgCtrlID & vbCr
                        strSndBuffer = strSndBuffer & EB & vbCr
                            
                        Call SendWSckData(strSndBuffer)
                    End If
                    
                Case "ORC"
                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mOrder.BarNo = strBarno
                    
                    If mOrder.MsgType = "ORM^O01" Then
                        Call GetOrder_BC6800(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                Case "PID"
                Case "PV1"
                Case "OBR"
                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mResult.BarNo = strBarno
                    
                   'Call SetSQLData("BAR", strBarno, "A")
                    
                    'If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = Trim(strBarno)
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    'End If
                    
                Case "OBX"
                    
                    'OBX|9|NM|6690-2^WBC^LN||5.90|10*3/uL|4.00-10.00|N|||F
                    'OBX|10|NM|704-7^BAS#^LN||0.09|10*3/uL|0.00-0.10|A|||F
                    'OBX|11|NM|706-2^BAS%^LN||1.5|%|0.0-1.0|H~A|||F
                    'OBX|12|NM|751-8^NEU#^LN||3.87|10*3/uL|2.00-7.00|A|||F
                
                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    If strIntBase = "" Then
                        strIntBase = Trim(mGetP(strRcvBuf, 5, "|"))
                    End If
                    
                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
                    
                    
                    'as-is : coulter
                    If strIntBase = "RBC" And IsNumeric(strResult) Then
                        strResult = strResult * 100
                    End If

                    If strIntBase = "PLT" And IsNumeric(strResult) Then
                        strResult = strResult / 10
                    End If
                
                    strIntResult = strResult
                    
                    'MsgBox strIntBase & "," & strResult
                    
                    'If Len(strIntBase) <= 5 Then
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    'End If
                    .spdResult.RowHeight(-1) = gROWHEIGHT
            End Select
        Next
        
        '(%WBC% * %NEUT%) / 100
        'Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
        
        '## DB에 결과저장
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)
            Call SetUpdateStatus(spdOrder, gRow, Res)
            strState = ""
        End If
    
    End With
    
    
    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_BC6800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

Public Sub TCPRcvData_BC5180()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim varResult       As Variant
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim intCnt          As Integer
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strTmp          As String
    
    Dim strTGResult     As String
    Dim strCHOLResult   As String
    Dim strHDLResult    As String
    Dim intCol          As Integer
    
    Dim blnResult       As Boolean
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strSndBuffer    As String
    
On Error GoTo ErrHandle
    
    strRecvData = Split(pBuffer, vbCr)
    
    With frmInterface
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    '오더요청
                    'MSH|^~\&||Mindray|||20081120174836||ORM^O01|4|P|2.3.1||||||UNICODE
                    
                    '검사결과(Sample)
                    'MSH|^~\&||Mindray|||20111124091140||ORU^R01|1|P|2.3.1||||||UNICODE
                    
                    '검사결과(QC)
                    'MSH|^~\&||Mindray|||20111124091422||ORU^R01|1|Q|2.3.1||||||UNICODE
                    
                    strOldBarno = ""
                    mOrder.MsgType = mGetP(strRcvBuf, 9, "|")
                    mOrder.MsgCtrlID = mGetP(strRcvBuf, 10, "|")
                    mOrder.TstType = mGetP(strRcvBuf, 11, "|")
                
                    If mOrder.MsgType = "ORU^R01" Then
                        '검사결과(Sample) Response Message
                        'MSH|^~\&|LIS||||20111124091140||ACK^R01|1|P|2.3.1||||||UNICODE
                        'MSA|AA|1
                        
                        '검사결과(QC) Response Message
                        'MSH|^~\&|LIS||||20111124091422||ACK^R01|1|Q|2.3.1||||||UNICODE
                        'MSA|AA|1
                                       strSndBuffer = SB
                        strSndBuffer = strSndBuffer & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|1|" & mOrder.TstType & "|2.3.1||||||UNICODE" & vbCr
                        strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MsgCtrlID & vbCr
                        strSndBuffer = strSndBuffer & EB & vbCr
                            
                        Call SendWSckData(strSndBuffer)
                    End If
                    
                Case "ORC"
                    'ORC|RF||SampleID1||IP
                    'ORC|RF||2012090228|BL
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mOrder.BarNo = strBarno
                    
                    If mOrder.MsgType = "ORM^O01" Then
                        Call GetOrder_BC6200(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                Case "PID"
                    'PID|1||^^^^MR

                Case "PV1"
                    'PV1|1
                    
                Case "OBR"
                    'OBR|1||ste5|00001^Automated Count^99MRC|||20111101170410|||||||||||||||||HM
                    'QC
                    'OBR|1||1|00003^LJ QCR^99MRC|||20201210085221|||||||||||||||||HM||||||||User
                    'Sample
                    'OBR|1||2012100197|00001^Automated Count^99MRC|||20201210113216|||||||||||||||||HM||||||||User

                    strBarno = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    mResult.BarNo = strBarno
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = Trim(strBarno)
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                Case "OBX"
                    'OBX|5|NM|6690-2^WBC^LN||6.58|10*9/L|4.00-10.00|N|||F
                    'QC
                    'OBX|5|NM|6690-2^WBC^LN||3.61|10*3/uL|3.10-4.10|N|||F
                    'Sample
                    'OBX|9|NM|6690-2^WBC^LN||4.57|10*3/uL|4.00-10.00|N|||F
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    If strIntBase = "" Then
                        strIntBase = Trim(mGetP(strRcvBuf, 5, "|"))
                    End If
                    
                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
                    strIntResult = strResult
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT
            End Select
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With
    
    
    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_BS240E" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

Public Sub TCPRcvData_BS360S()
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim varResult       As Variant
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim intCnt          As Integer
    
    Dim strOrderCode    As String   '처방코드
    Dim strTestCode     As String   '검사코드
    Dim strTestSubCode  As String   '검사코드
    Dim strTestName     As String   '검사명
    Dim strSeqNo        As String   '로컬DB 검사Seq
    
    Dim strTmp          As String
    
    Dim strTGResult     As String
    Dim strCHOLResult   As String
    Dim strHDLResult    As String
    Dim intCol          As Integer
    
    Dim blnResult       As Boolean
    
    Dim strRstRow       As String   '결과스프레드 현재 Row
    Dim strDecYN        As String   '결과판정여부
    Dim strJudge        As String   '결과판정
    
    Dim strQCData       As String
    Dim i               As Integer
    Dim Res             As Integer
    Dim strSndBuffer    As String
    
On Error GoTo ErrHandle
    
    strRecvData = Split(pBuffer, vbCr)
    
    With frmInterface
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    mOrder.BSMaker = mGetP(strRcvBuf, 3, "|")
                    mOrder.BSMchNm = mGetP(strRcvBuf, 4, "|")
                    mOrder.BSMType = mGetP(strRcvBuf, 9, "|")
                    mOrder.BSDtTm = Format(Now, "yyyymmddhhmmss")
                
                    With mOrder
                        .MSHCorpName = mGetP(strRcvBuf, 3, "|")
                        .MSHDeviceModel = mGetP(strRcvBuf, 4, "|")
                        .MSHSysDateTime = mGetP(strRcvBuf, 7, "|")
                        .MSHMessageType = mGetP(strRcvBuf, 9, "|")
                        .MSHMessageID = mGetP(strRcvBuf, 10, "|")
                        .MSHProduct = mGetP(strRcvBuf, 11, "|")
                        .MSHHL7Version = mGetP(strRcvBuf, 12, "|")
                        .MSHResultType = mGetP(strRcvBuf, 16, "|")
                        .MSHChrEncoding = mGetP(strRcvBuf, 18, "|")
                    End With
                    
                    Select Case mOrder.MSHMessageType
                        '-- 검사결과수신 ACK
                        Case "ORU^R01"  '==> ACK^R01
                                           strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
                            strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
                            strSndBuffer = strSndBuffer & EB & vbCr
                            
                            Call SendWSckData(strSndBuffer)
                            
                        '-- 오더요청수신
                        Case "QRY^Q02"  '==> QCK^Q02
                            strSndBuffer = ""
                            With spdOrder
                                For i = 1 To .MaxRows
                                    If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
                                        '-- 오더있음
                                                       strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||QCK^Q02|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
                                        strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
                                        strSndBuffer = strSndBuffer & "ERR|0|" & vbCr & EB & vbCr
                                        strSndBuffer = strSndBuffer & "QAK|SR|OK|" & vbCr
                                        strSndBuffer = strSndBuffer & EB & vbCr

                                        Call SendWSckData(strSndBuffer)
                                        Exit For
                                    End If
                                Next
                            End With

                            '-- 오더없음
                            If strSndBuffer = "" Then
                                               strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||QCK^Q02|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
                                strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
                                strSndBuffer = strSndBuffer & "ERR|0|" & vbCr & EB & vbCr
                                strSndBuffer = strSndBuffer & "QAK|SR|NF|" & vbCr
                                strSndBuffer = strSndBuffer & EB & vbCr

                                Call SendWSckData(strSndBuffer)
                            End If
                        
                        '-- 오더 전송
                        Case "ACK^Q03"
                            '-- 최초이후전송
                            Call GetOrder_BS360S(strBarno, gHOSP.RSTTYPE)
                            
                    End Select
                        
                Case "QRD"
                    With mOrder
                        .QRDQryTime = mGetP(strRcvBuf, 2, "|")
                        .QRDQryFormatCode = mGetP(strRcvBuf, 3, "|")
                        .QRDQryPriority = mGetP(strRcvBuf, 4, "|")
                        .QRDNum = mGetP(strRcvBuf, 5, "|")
                        .QRDQLRequest = mGetP(strRcvBuf, 8, "|")
                        .QRDSampleBarcode = mGetP(strRcvBuf, 9, "|")
                        .QRDWSFilter = mGetP(strRcvBuf, 10, "|")
                        .QRDQryResultLevel = mGetP(strRcvBuf, 13, "|")
                        
                        'mOrder.Seq = mGetP(strRcvBuf, 5, "|")
                        'mOrder.BarNo = mOrder.Seq
                        'strBarno = mOrder.Seq
                    End With
                    
                Case "QRF"
                    mOrder.BSModel = mGetP(strRcvBuf, 2, "|")
                    mOrder.BSSTime = mGetP(strRcvBuf, 3, "|")
                    mOrder.BSETime = mGetP(strRcvBuf, 4, "|")
                    mOrder.BSQRF = strRcvBuf
                    mOrder.Seq = mGetP(strRcvBuf, 5, "|")
                                        
                    With mOrder
                        .QRFProduct = mGetP(strRcvBuf, 2, "|")
                        .QRFWherStartDtTm = mGetP(strRcvBuf, 3, "|")
                        .QRFWherEndDtTm = mGetP(strRcvBuf, 4, "|")
                        .QRFWhichDtTmQualifier = mGetP(strRcvBuf, 7, "|")
                        .QRFWhichStatusQualifier = mGetP(strRcvBuf, 8, "|")
                        .QRFDtTmSelecQualifier = mGetP(strRcvBuf, 9, "|")
                    End With
                    
                    '-- 최초오더전송
                    intSndPhase = 1
                    
                    Call GetOrder_BS360S(strBarno, gHOSP.RSTTYPE)
                    
                Case "PID"

                Case "PV1"
                    
                Case "OBR"
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    mResult.BarNo = strBarno
                    
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = Trim(strBarno)
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                Case "OBX"
                    strIntBase = mGetP(strRcvBuf, 4, "|")
                    If strIntBase = "" Then
                        strIntBase = Trim(mGetP(strRcvBuf, 5, "|"))
                    End If
                    
                    strResult = Trim$(mGetP(strRcvBuf, 6, "|"))
                    strIntResult = strResult
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT
            End Select
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With
    
    
    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_BS360S" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

Private Sub GetOrder_BC6200(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim strSend     As String
    Dim strSndMsg   As String
    
On Error GoTo ErrHandle

    intRow = -1
    strSend = ""

    ''Call SetCommStatus("Q", pBarNo, frmInterface.spdComStatus)
    ''Call SetCommStatus("Q", pBarNo, frmInterface.lstComStatus)
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_BC6200(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
                      
                      strSend = SB
            strSend = strSend & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ORR^O02|1|P|2.3.1||||||UNICODE" & vbCr
            strSend = strSend & "MSA|AR|" & mOrder.MsgCtrlID & vbCr
            strSend = strSend & EB & vbCr

            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
            
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 오더문자열만들기
                      strSend = SB
            strSend = strSend & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ORR^O02|1|P|2.3.1||||||UNICODE" & vbCr
            strSend = strSend & "MSA|AA|" & mOrder.MsgCtrlID & vbCr
            strSend = strSend & "PID|1||" & mOrder.PID & "^^^^MR||^" & mOrder.PNAME & "|||" & vbCr
            strSend = strSend & "PV1|1|E|^^||||||||||||||||||AF|" & mOrder.BarNo & "|||" & vbCr
            strSend = strSend & "ORC|AF|" & mOrder.BarNo & "|||" & vbCr
            strSend = strSend & "OBR|1|" & mOrder.BarNo & "||||||||||||||||||||||HM||||||||" & vbCr
            strSend = strSend & "OBX|1|IS|08001^Take Mode^99MRC||A||||||F" & vbCr                   'O : open-vial,     A : autoloading,    C : closed-tube
            strSend = strSend & "OBX|2|IS|08002^Blood Mode^99MRC||W||||||F" & vbCr                  'W : whole blood,   P : predilute,      B : body fluid,     Q : control,    M : micro-WB
            strSend = strSend & "OBX|3|IS|08003^Test Mode^99MRC||" & strItems & "||||||F" & vbCr    'CBC, CBC+DIFF, CBC+RET, CBC+NRBC, CBC+DIFF+RET, CBC+DIFF+NRBC, CBC+DIFF+RET+NRBC, RET
            strSend = strSend & "OBX|4|IS|01002^Ref Group^99MRC||||||||F" & vbCr
            strSend = strSend & "OBX|5|NM|30525-0^Age^LN|||hr|||||F" & vbCr
            strSend = strSend & "OBX|6|ST|01001^Remark^99MRC||||||||F" & vbCr
            strSend = strSend & EB & vbCr
            
'BC6800
'MSH|^~\&|LIS|Mindray|BC-6800||20201210082352||ORR^O02||P|2.31||||||UNICODE
'MSA|AA|2
'PID|1||08012291^^^^MR||^김민자|||Male
'PV1|1||^^|||||||||||||||||
'ORC|AF||2001311551
'OBR|1|2001311551||00001^Automated Count^99MRC||20201210082352||||||||20201210082352||||||||||HM||||||||admin
'OBX|1|IS|08001^Take Mode^99MRC||A||||||F
'OBX|2|IS|08002^Blood Mode^99MRC||W||||||F
'OBX|3|IS|08003^Test Mode^99MRC||CBC+DIFF||||||F
'OBX|2|IS|01002^Ref Group^99MRC||General||||||F
'OBX|3|NM|30525-0^Age^LN|||yr|||||F
'OBX|4|ST|01001^Remark^99MRC||||||||F
'OBX|5|ST|08005^SerialNumber^99MRC||||||||F
'OBX|6|IS|01007^Sample Type^99MRC||Venous blood||||||F
'OBX|7|IS|01008^Patient Area^99MRC||||||||F
'OBX|8|ST|01009^Custom patient info 1^99MRC||Nothing||||||F
'OBX|9|ST|01010^Custom patient info 2^99MRC||Nothing||||||F
'OBX|10|ST|01011^Custom patient info 3^99MRC||Nothing||||||F

            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "0", intRow, colCHECKBOX)
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
            
            Call SendWSckData(strSend)
            
            intSndPhase = intSndPhase + 1
        End If

        '-- 현재 Row
        gRow = intRow

    End With

    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetOrder_BC6200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show


End Sub

Private Sub GetOrder_BC6800(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim strSend     As String
    Dim strSndMsg   As String
    
On Error GoTo ErrHandle

    intRow = -1
    strSend = ""
    
    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_BC6800(gHOSP.MACHCD, pBarno, intRow)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
                      
                      strSend = SB
            strSend = strSend & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ORR^O02|1|P|2.3.1||||||UNICODE" & vbCr
            strSend = strSend & "MSA|AR|" & mOrder.MsgCtrlID & vbCr
            strSend = strSend & EB & vbCr

            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "오더없음", intRow, colSTATE)
            
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 오더문자열만들기
                      strSend = SB
            strSend = strSend & "MSH|^~\&|LIS||||" & Format(Now, "yyyymmddhhmmss") & "||ORR^O02|1|P|2.3.1||||||UNICODE" & vbCr
            strSend = strSend & "MSA|AA|" & mOrder.MsgCtrlID & vbCr
            strSend = strSend & "PID|1||" & mOrder.PID & "^^^^MR||^" & mOrder.PNAME & "|||" & vbCr
            strSend = strSend & "PV1|1|E|^^||||||||||||||||||AF|" & mOrder.BarNo & "|||" & vbCr
            strSend = strSend & "ORC|AF|" & mOrder.BarNo & "|||" & vbCr
            strSend = strSend & "OBR|1|" & mOrder.BarNo & "||||||||||||||||||||||HM||||||||" & vbCr
            strSend = strSend & "OBX|1|IS|08001^Take Mode^99MRC||A||||||F" & vbCr                   'O : open-vial,     A : autoloading,    C : closed-tube
            strSend = strSend & "OBX|2|IS|08002^Blood Mode^99MRC||W||||||F" & vbCr                  'W : whole blood,   P : predilute,      B : body fluid,     Q : control,    M : micro-WB
            strSend = strSend & "OBX|3|IS|08003^Test Mode^99MRC||" & strItems & "||||||F" & vbCr    'CBC, CBC+DIFF, CBC+RET, CBC+NRBC, CBC+DIFF+RET, CBC+DIFF+NRBC, CBC+DIFF+RET+NRBC, RET
            strSend = strSend & "OBX|4|IS|01002^Ref Group^99MRC||||||||F" & vbCr
            strSend = strSend & "OBX|5|NM|30525-0^Age^LN|||hr|||||F" & vbCr
            strSend = strSend & "OBX|6|ST|01001^Remark^99MRC||||||||F" & vbCr
            strSend = strSend & EB & vbCr
            
'BC6800
'MSH|^~\&|LIS|Mindray|BC-6800||20201210082352||ORR^O02||P|2.31||||||UNICODE
'MSA|AA|2
'PID|1||08012291^^^^MR||^김민자|||Male
'PV1|1||^^|||||||||||||||||
'ORC|AF||2001311551
'OBR|1|2001311551||00001^Automated Count^99MRC||20201210082352||||||||20201210082352||||||||||HM||||||||admin
'OBX|1|IS|08001^Take Mode^99MRC||A||||||F
'OBX|2|IS|08002^Blood Mode^99MRC||W||||||F
'OBX|3|IS|08003^Test Mode^99MRC||CBC+DIFF||||||F
'OBX|2|IS|01002^Ref Group^99MRC||General||||||F
'OBX|3|NM|30525-0^Age^LN|||yr|||||F
'OBX|4|ST|01001^Remark^99MRC||||||||F
'OBX|5|ST|08005^SerialNumber^99MRC||||||||F
'OBX|6|IS|01007^Sample Type^99MRC||Venous blood||||||F
'OBX|7|IS|01008^Patient Area^99MRC||||||||F
'OBX|8|ST|01009^Custom patient info 1^99MRC||Nothing||||||F
'OBX|9|ST|01010^Custom patient info 2^99MRC||Nothing||||||F
'OBX|10|ST|01011^Custom patient info 3^99MRC||Nothing||||||F

            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "0", intRow, colCHECKBOX)
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
            
            Call SendWSckData(strSend)
            
            intSndPhase = intSndPhase + 1
        End If

        '-- 현재 Row
        gRow = intRow

    End With

    Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetOrder_BC6800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show


End Sub



'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder_BS360S(ByVal pBarno As String, ByVal pType As String) ', _
                               ByVal pMaker As String, ByVal pMchNm As String, _
                               ByVal pModel As String, ByVal pSTime As String, _
                               ByVal pETime As String, ByVal pQryId As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    Dim strSend     As String
    Dim blnLast     As Boolean
    Dim strNow      As String
    
    intRow = -1

    '-- 1. 접수정보 조회
    With frmInterface
        Select Case pType
            '-- 바코드 사용
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarno Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(spdOrder, i, colCHECKBOX) = "1" Then
                        'pBarno = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        'mOrder.BarNo = pBarno
                        'intRow = i
                        
                        pBarno = Trim(GetText(.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarno
                        mOrder.PNAME = Trim(GetText(.spdOrder, i, colPNAME))
                        mOrder.PID = Trim(GetText(.spdOrder, i, colPID))
                        mOrder.ChartNo = Trim(GetText(.spdOrder, i, colCHARTNO))
                        
                        
                        mOrder.Seq = Trim(GetText(.spdOrder, i, colSEQNO))
                        
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            If UCase(gHOSP.MACHNM) = "BS360S" Then
               Exit Sub
            Else
                intRow = .spdOrder.DataRowCnt + 1
                If .spdOrder.MaxRows < intRow Then
                    .spdOrder.MaxRows = intRow
                End If
            End If
        End If

        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0

        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = gROWHEIGHT
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = ""
        strItems = GetEquipExamCode_BS360S(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            '-- 진행상태(Order) 표시
            Call SetText(spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 진행상태(Order) 표시
            Call SetText(spdOrder, "오더준비", intRow, colSTATE)
            
'''            strNow = Format(Now, "yyyymmddhhmmss")
'''
'''            '-- 오더문자열만들기
'''                      strSend = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||DSR^Q03|1|P|2.3.1||||0||ASCII|||" & vbCr
'''            strSend = strSend & "MSA|AA|" & CStr(intSndPhase) & "|Message accepted|||0|" & vbCr
'''            strSend = strSend & "ERR|0|" & vbCr
'''            strSend = strSend & "QAK|SR|OK|" & vbCr
'''
'''            strSend = strSend & "QRD|" & strNow & "|R|D|1|||RCT|COR|ALL||" & vbCr
'''            'strSend = strSend & "QRD|" & strNow & "|R|D|" & mOrder.BSQryId & "|||RCT|COR|ALL||" & vbCr
'''
'''            strSend = strSend & mOrder.BSQRF & vbCr
'''            'strSend = strSend & "QRD|BS-360S|" & strNow & "|" & strNow & "|||RD||OTH|||T|" & vbCr
'''
'''            strSend = strSend & "DSP|1||" & mOrder.ChartNo & "|||" & vbCr   'Patient ID
'''            strSend = strSend & "DSP|2|||||" & vbCr                         'Bed No
'''            strSend = strSend & "DSP|3||" & mOrder.PNAME & "|||" & vbCr     'Patient Name
'''            strSend = strSend & "DSP|4|||||" & vbCr                         'Birth Date
'''            strSend = strSend & "DSP|5|||||" & vbCr                         'Gender
'''            strSend = strSend & "DSP|6|||||" & vbCr                         'Blood Type
'''            strSend = strSend & "DSP|7|||||" & vbCr
'''            strSend = strSend & "DSP|8|||||" & vbCr
'''            strSend = strSend & "DSP|9|||||" & vbCr
'''            strSend = strSend & "DSP|10|||||" & vbCr
'''            strSend = strSend & "DSP|11|||||" & vbCr
'''            strSend = strSend & "DSP|12|||||" & vbCr
'''            strSend = strSend & "DSP|13|||||" & vbCr
'''            strSend = strSend & "DSP|14|||||" & vbCr
'''            strSend = strSend & "DSP|15|||||" & vbCr                        'Patient Type
'''            strSend = strSend & "DSP|16|||||" & vbCr
'''            strSend = strSend & "DSP|17|||||" & vbCr                        'Pay type
'''            strSend = strSend & "DSP|18|||||" & vbCr
'''            strSend = strSend & "DSP|19|||||" & vbCr
'''            strSend = strSend & "DSP|20|||||" & vbCr
'''            strSend = strSend & "DSP|21||" & mOrder.BarNo & "|||" & vbCr    'barcode
'''            strSend = strSend & "DSP|22||" & mOrder.ChartNo & "|||" & vbCr  'sample id
'''            strSend = strSend & "DSP|23||" & strNow & "|||" & vbCr          'send time
'''            strSend = strSend & "DSP|24||N|||" & vbCr                       'STAT (응급여부)
'''            strSend = strSend & "DSP|25|||||" & vbCr
'''            strSend = strSend & "DSP|26||serum|||" & vbCr                   'sample type
'''            strSend = strSend & "DSP|27|||||" & vbCr                        'doctor(처방의)
'''            strSend = strSend & "DSP|28|||||" & vbCr                        'send department
'''            strSend = strSend & strItems
'''
'''            blnLast = False
'''            For i = 1 To spdOrder.DataRowCnt
'''                If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
'''                    blnLast = True
'''                    Exit For
'''                End If
'''            Next i
'''
'''            If blnLast = True Then
'''                strSend = strSend & "DSC|" & CStr(intSndPhase) & "|" & vbCr
'''            Else
'''                strSend = strSend & "DSC||" & vbCr
'''            End If
'''
'''            strSend = strSend & EB & vbCr
'''
'''            '-- 진행상태(Order) 표시
'''            Call SetText(frmInterface.spdOrder, "", intRow, colCHECKBOX)
'''            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
'''            Call SendWSckData(strSend)
            
            
            '-- 오더문자열만들기
                      strSend = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & mOrder.BSDtTm & "||DSR^Q03|1|P|2.3.1||||0||ASCII|||" & vbCr
            strSend = strSend & "MSA|AA|" & CStr(intSndPhase) & "|Message accepted|||0|" & vbCr
            strSend = strSend & "ERR|0|" & vbCr
            strSend = strSend & "QAK|SR|OK|" & vbCr
            strSend = strSend & "QRD|" & mOrder.BSDtTm & "|R|D|" & mOrder.BSQryId & "|||RD||OTH|||T|" & vbCr
            strSend = strSend & mOrder.BSQRF & vbCr
            strSend = strSend & "DSP|1||" & mOrder.ChartNo & "|||" & vbCr       'Patient ID
            strSend = strSend & "DSP|2|||||" & vbCr                         'Bed No
            strSend = strSend & "DSP|3||" & mOrder.PNAME & "|||" & vbCr     'Patient Name
            strSend = strSend & "DSP|4|||||" & vbCr                         'Birth Date
            strSend = strSend & "DSP|5|||||" & vbCr                         'Gender
            strSend = strSend & "DSP|6|||||" & vbCr                         'Blood Type
            strSend = strSend & "DSP|7|||||" & vbCr
            strSend = strSend & "DSP|8|||||" & vbCr
            strSend = strSend & "DSP|9|||||" & vbCr
            strSend = strSend & "DSP|10|||||" & vbCr
            strSend = strSend & "DSP|11|||||" & vbCr
            strSend = strSend & "DSP|12|||||" & vbCr
            strSend = strSend & "DSP|13|||||" & vbCr
            strSend = strSend & "DSP|14|||||" & vbCr
            strSend = strSend & "DSP|15|||||" & vbCr                        'Patient Type
            strSend = strSend & "DSP|16|||||" & vbCr
            strSend = strSend & "DSP|17|||||" & vbCr                        'Pay type
            strSend = strSend & "DSP|18|||||" & vbCr
            strSend = strSend & "DSP|19|||||" & vbCr
            strSend = strSend & "DSP|20|||||" & vbCr
            '바코드로 사용할때만 필요
            'If gHOSP.BARUSE = "Y" Then
                strSend = strSend & "DSP|21||" & mOrder.BarNo & "|||" & vbCr      'barcode
            'Else
            '    strSend = strSend & "DSP|21|||||" & vbCr                        'barcode
            'End If
            'strSend = strSend & "DSP|22||" & CStr(mOrder.Seq) & "|||" & vbCr                'sample id
            
            '2020-07-02 수정
            'strSend = strSend & "DSP|22||" & CStr(intSndPhase) & "|||" & vbCr               'sample id
            'strSend = strSend & "DSP|22||" & mOrder.ChartNo & "|||" & vbCr               'sample id
            strSend = strSend & "DSP|22||" & mOrder.Seq & "|||" & vbCr                'sample id
            
            
            strSend = strSend & "DSP|23||" & Format(Now, "yyyymmddhhmmss") & "|||" & vbCr   'send time
            strSend = strSend & "DSP|24||N|||" & vbCr                                       'STAT (응급여부)
            strSend = strSend & "DSP|25|||||" & vbCr
            strSend = strSend & "DSP|26||serum|||" & vbCr                                   'sample type
            strSend = strSend & "DSP|27|||||" & vbCr                                        'doctor(처방의)
            strSend = strSend & "DSP|28|||||" & vbCr                                        'send department
            strSend = strSend & strItems
            
            blnLast = False
            For i = 1 To spdOrder.DataRowCnt
                If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
                    blnLast = True
                    Exit For
                End If
            Next i
            
            If blnLast = True Then
                strSend = strSend & "DSC|" & CStr(intSndPhase) & "|" & vbCr
            Else
                strSend = strSend & "DSC||" & vbCr
            End If
            
            strSend = strSend & EB & vbCr
            
            '-- 진행상태(Order) 표시
            Call SetText(frmInterface.spdOrder, "", intRow, colCHECKBOX)
            Call SetText(frmInterface.spdOrder, "오더전송", intRow, colSTATE)
            
            
            SetRawData "[Tx]" & strSend
            wSck.SendData strSend
            intSndPhase = intSndPhase + 1
            
        End If

        '-- 현재 Row
        gRow = intRow

    End With

End Sub


Private Sub TCPRcvData_PPC300N_OLD()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strTypeSeq      As String   '수신한 Record Type Seq
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    Dim strTemp         As String
    Dim strSend         As String
    Dim strOrder        As String
    Dim strLot          As String
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strTemp = mGetP(strRcvBuf, 4, "|")
            strType = mGetP(strTemp, 1, ";")
            strTypeSeq = mGetP(strTemp, 2, ";")
            
            Select Case strType
                Case "REQ"
                    If strTypeSeq = "1" Then
                        '오더요청 REQ;1
                        'Request information
                        'Start time         2010/11/01 00:00:00
                        'End time           2010/11/01 23:59:59
                        '<SB>|;^\|U8030|REQ;1|2010/11/01^00:00:00;2010/11/01^23:59:59|ASCII|<EB>

                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
''                        strSend = SB & "|;^\|LisDemo|" & "ASW;1||ASCII|" & EB
''                        SetRawData "[Tx]" & strSend
''                        wSck.SendData strSend
''                        '3.LIS Busy : <SB> |;^\|LISDEMO|ASW;1||ASCII|<EB>
''
''                        strSend = SB & "|;^\|LisDemo|" & "ASW;3||ASCII|" & EB
''                        SetRawData "[Tx]" & strSend
''                        wSck.SendData strSend
''                        '3.No Order : <SB> |;^\|LISDEMO|ASW;3||ASCII|<EB>
                        
                        strOrder = "" '5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP
                        intOrdCnt = 0
                        With spdOrder
                            For i = 1 To .MaxRows
                                .Row = i
                                .Col = colCHECKBOX
                                If .Value = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
                                    intOrdCnt = intOrdCnt + 1
                                    'strOrder = strOrder & GetText(spdOrder, i, colBARCODE) & ";" & GetText(spdOrder, i, colDEPT) & ";"
                                    
                                    strOrder = strOrder & GetText(spdOrder, i, colBARCODE) & ";"
                                    strOrder = strOrder & GetTag(spdOrder, i, colSTATE) & ";"
                                    
                                    Call SetText(spdOrder, "0", i, colCHECKBOX)
                                    Call SetText(spdOrder, "오더전송", i, colSTATE)
                                End If
                            Next
                        End With
                        
                        If strOrder = "" And intOrdCnt = 0 Then
                            strSend = SB & "|;^\|LisDemo|" & "ASW;3||ASCII|" & EB
                            SetRawData "[Tx]" & strSend
                            wSck.SendData strSend
                            '3.No Order : <SB> |;^\|LISDEMO|ASW;3||ASCII|<EB>
                        Else
                            strOrder = Mid(strOrder, 1, Len(strOrder) - 1)
                            strOrder = CStr(intOrdCnt) & ";" & strOrder
                            
                            strSend = SB & "|;^\|LisDemo|TRA;5|" & strOrder & "|ASCII|" & EB
                            SetRawData "[Tx]" & strSend
                            wSck.SendData strSend
                            '3.Order : SB & "|;^\|LisDemo|TRA;5|5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP|ASCII|" & EB
                        End If
                        
                        strState = "Q"
                        '<SB>|;^\|LisDemo|TRA;5|5;12345678;AST^ALT^TP^GLU_HK;23456789;TP;34567890;ALT;45678901;TP^DB;56789012;AST^ALT^TP^GLU_HK^ALP|ASCII|<EB>
                    Else
                        '일반샘플 REQ;2
                        '샘플정보 REQ;3
                        'QC  정보 REQ;4
                        'Cal 정보 REQ;5
                        
                        'Request transferring results
                        '1.RCV  : <SB>|;^\|U8030|REQ;2|1234;2|ASCII|<EB>
                        strTemp = mGetP(strRcvBuf, 5, "|")
                        strBarno = Trim(mGetP(strTemp, 1, ";"))     'BarCode
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
                        strSend = SB & "|;^\|LisDemo|" & "ASW;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB> |;^\|LISDEMO|ASW;2||ASCII|<EB>
                    
                        If Trim(strBarno) <> Trim(strOldBarno) Then
                            strOldBarno = strBarno
                            '-- 결과정보
                            With mResult
                                .BarNo = strBarno
                                .RsltDate = Format(Now, "yyyy-mm-dd")
                                .RsltTime = Format(Now, "hh:mm:ss")
                                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            End With
                        End If
                        
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                        If gRow <= 0 Then
                            Exit Sub
                        End If
                        
                        strState = "O"
                        
                    End If
                    
                Case "ASK"
                    '--
                
                Case "TRA"
                    '1.RCV  : <SB>|;^\|U8030|TRA;2|1;201009200001;1234;;ALT;;43;;U/L;0;40;;;|ASCII|<EB>
                    strTemp = mGetP(strRcvBuf, 5, "|")
                    
                    '샘플정보
                    If strTypeSeq = "1" Then
                        '샘플정보 TRA;1
                        strSeq = mGetP(strTemp, 1, ";")
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;" & strSeq & "||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                    
                    '일반결과정보
                    ElseIf strTypeSeq = "2" Then
                        strIntBase = Trim(mGetP(strTemp, 5, ";"))   'Item
                        strResult = Trim(mGetP(strTemp, 7, ";"))    'Result
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;6|" & strBarno & ";" & strIntBase & "|ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                        
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                    'QC 정보
                    ElseIf strTypeSeq = "3" Then
                        strIntBase = Trim(mGetP(strTemp, 1, ";"))   'Item
                        strLot = Trim(mGetP(strTemp, 4, ";"))   'Lot
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;" & strSeq & "||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                        
                    'Cal 정보
                    ElseIf strTypeSeq = "4" Then
                        strIntBase = Trim(mGetP(strTemp, 1, ";"))   'Item
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;7" & "|" & strLot & ";" & strIntBase & "|ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                    End If
                    
                    .spdResult.RowHeight(-1) = gROWHEIGHT
                
                Case "END"
                    '1.RCV  : <SB>|;^\|U8030|END;1||ASCII|<EB>
                    
                    strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                    SetRawData "[Tx]" & strSend
                    wSck.SendData strSend
                    '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                    
                    strSend = SB & "|;^\|LisDemo|" & "REP;2||ASCII|" & EB
                    SetRawData "[Tx]" & strSend
                    wSck.SendData strSend
                    '2.SEND : <SB>|;^\|LisDemo|REP;2||ASCII| <EB>
            
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
            
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If

            End Select
        
        Next
    
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_F200()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- 인터페이스 응답
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||ASCII|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|TR03-025|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- 참고치
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- 남자참고치를 기본으로 한다
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- 결과Row 추가
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- 소수점 처리
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- 결과판정
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- 진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '-- 메인화면 결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
                            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '순번
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '장비채널
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS결과
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
                            
                            '-- 이전결과 조회
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '이전결과
                            
                            '-- H/L 색깔표시
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- 로컬 저장
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = gROWHEIGHT
        
                    End If

                    .spdResult.RowHeight(-1) = gROWHEIGHT

            End Select
        Next
    
        '## DB에 결과저장
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HITACHI7180()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim strGFR          As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strCRP          As String
    Dim strRF           As String
    
    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    
'    Dim strGA           As String
'    Dim strGAAlb        As String
    
'On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY 수신
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- 장비로 전송
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
                    ';A1     0 861   1000001319    0091719173849

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    sFunc = Mid(strRcvBuf, 2, 40)
                    
                    With mOrder
                        .BarNo = strBarno
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                    End With
                    
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration 데이터는 무시함
'                    If UCase(strFunc) = "K" Or UCase(strFunc) = "L" Or UCase(strFunc) = "G" Or UCase(strFunc) = "H" Then
'                        '-- 장비로 전송
'                        Call SendData(SndMore)
'                        strState = ""
'                        Exit Sub
'                    End If
                    
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
'''
'''                    If UCase(strFunc) = "F" Then
'''                        '-- 장비로 전송
'''                        Call SendData(SndMore)
'''                        strState = ""
'''                        Exit Sub
'''                    End If
                    
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If
                    
                    strTC = ""
                    strTG = ""
                    strHDL = ""
                    
                    For ii = 51 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                        strIntResult = strResult
                        strComm = Trim(Mid(strRcvBuf, ii + 9, 1))
            
                        '-- CREA 결과저장
'                        If Trim(strIntBase) = "2" Then
'                            strGFR = ""
'                            strResult = Format(strResult, "##0.00")
'                            strCREA = strResult
'
'                            If mPatient.AGE <> "" And mPatient.SEX <> "" Then
'                                If CCur(strResult) > 0 Then
'                                    '18세 이상만 적용
'                                    If IsNumeric(strCREA) And mPatient.AGE > 18 Then
'                                        If mPatient.SEX = "M" Then
'                                            strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203)
'                                        ElseIf mPatient.SEX = "F" Then
'                                            strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
'                                        End If
'
'                                        If strGFR <> "" Then
'                                            strGFR = Format(strGFR, "##0.00")
'                                            If strGFR <= 120 Then
'                                                strGFR = Round(strGFR, 2)
'                                            ElseIf strGFR > 120 Then
'                                                strGFR = "> 120"
'                                            End If
'                                        End If
'                                    End If
'                                Else
'                                    strGFR = "Error"
'                                End If
'                            End If
'                        End If
'
'                        If strIntBase = "20" Then    'CRP
'                            strCRP = strResult
'                        End If
'
'                        If strIntBase = "21" Then    'RF
'                            strRF = strResult
'                        End If
'
'                        If strIntBase = "19" Then    'TCHO
'                            strTC = strResult
'                        End If
'
'                        If strIntBase = "10" Then   'TG
'                            strTG = strResult
'                        End If
'
'                        If strIntBase = "9" Then    'HDLC
'                            strHDL = strResult
'                        End If
                    
'                        If strIntBase = "25" Then    'GA
'                            strGA = strResult
'                        End If
'
'                        If strIntBase = "26" Then    'GA-Alb
'                            strGAAlb = strResult
'                        End If
                    
ReCal:
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                    Next
                    
                    'GA% = GA + (ALB - BCP)
                    '장비에서 결과나옴
'                    If strGA <> "" And strGAAlb <> "" And IsNumeric(strGA) And IsNumeric(strGAAlb) Then
'                        strIntBase = "77"
'                        strResult = strGA + (strGAAlb - strHDL)
'                        If strResult < 0 Then
'                            strResult = "0"
'                        End If
'                        strIntResult = ""
'                        strGA = ""
'                        strGAAlb = ""
'                        'strHDL = ""
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'
'                    End If
                    
                    'LDL 계산
'                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
'                        strIntBase = "95"
'                        strResult = strTC - ((strTG / 5) + strHDL)
'                        If strResult < 0 Then
'                            strResult = "0"
'                        End If
'                        strIntResult = ""
'                        strTC = ""
'                        strTG = ""
'                        strHDL = ""
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'
'                    End If
'
'                    'CRP 정성
'                    If strCRP <> "" Then
'                        strIntBase = "87"
'                        If strCRP < 0.5 Then
'                            strResult = "Negative (" & strCRP & ")"
'                        Else
'                            strResult = "Positive (" & strCRP & ")"
'                        End If
'                        strIntResult = ""
'                        strCRP = ""
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'                    End If
'
'                    'RA Fact 정성
'                    If strRF <> "" Then
'                        strIntBase = "88"
'                        If strRF < 15 Then
'                            strResult = "Negative (" & strRF & ")"
'                        Else
'                            strResult = "Positive (" & strRF & ")"
'                        End If
'                        strIntResult = ""
'                        strRF = ""
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'                    End If
'
'                    '-- GFR 저장
'                    If strGFR <> "" Then
'                        strIntBase = "89"
'                        strResult = strGFR
'                        strIntResult = ""
'                        strGFR = ""
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'                    End If
    
                    Call SendData(SndMore)
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SerialRcvData_H7180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        'RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function


Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub


Private Sub SendMessage_1200(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim r As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & R_S _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & R_S _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & R_S & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & R_S & ETX
            
        Case "SMP_ORD"
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    comEqp.Output = sSendData
    
SendMessage_Error:
    If Err <> 0 Then
'        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub SerialRcvData_RP500()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
        
        
    Dim strRcvMsg2      As String
    Dim strRcvMsg3      As String
    Dim strRcvMsg7      As String
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    Dim r   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim Result  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
On Error GoTo Err
    
    
    X = InStr(1, RcvBuffer, FS)
    If RcvBuffer <> "" Then
        MsgID = Mid(RcvBuffer, 2, X - 2)
    End If
    Select Case MsgID
        Case "ID_REQ"
            Call SendMessage_1200("ID_DATA")
        Case "SMP_START"
        Case "SMP_NEW_AV"
            Do Until X = 0
                X = InStr(X, RcvBuffer, "r")
                If X = 0 Then Exit Do
                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                    X = X + 5
                    C = InStr(X, RcvBuffer, GS)
                    Sample_Seq = Mid(RcvBuffer, X, C - X)
                End If
                Call GetaModiIID(RcvBuffer)
                Call SendMessage_1200("SMP_REQ")
            Loop
        
        Case "SYS_READY"
        Case "SYS_NOT_READY"
        Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
            GoTo RST
        Case "CAL_ABORT"
    End Select
    
    Exit Sub

RST:

    MsgBuf = RcvBuffer
    
    
    With frmInterface
        If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
            'aMod
            x1 = 1
            x1 = InStr(x1, MsgBuf, "aMod") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                aMod = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'iIID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iIID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iIID = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'rSEQ
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                rSeq = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'PID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iPID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iPID = Mid(MsgBuf, x1, x2 - x1)
            End If
            'DATE
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rDATE") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstDate = Mid(MsgBuf, x1, x2 - x1)
                sRstDate = ConvertDateType(sRstDate)
            End If
            'TIME
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rTIME") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstTime = Mid(MsgBuf, x1, x2 - x1)
                sRstTime = Format(sRstTime, "HHNNSS")
            End If
        
            x2 = 0
        
            '접수번호, SeqNo
            strBarno = Trim(iPID)
            strSeqNo = Trim(rSeq)
            
            If Trim(strBarno) = "" Then Exit Sub
            
            '-- 결과정보
            With mResult
                .BarNo = strBarno
                .Seq = strSeqNo
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
            
            '-- 결과환자정보
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
        
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                
                strResult = Mid(MsgBuf, x2, x1 - x2)
                strIntResult = strResult
                
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
            Loop
            
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "c") <> 0
                x1 = InStr(x1, MsgBuf, FS & "c")
                x2 = InStr(x1, MsgBuf, GS)
    
        '        'AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
    
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
                strIntResult = strResult
            
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
            Loop
            
            spdOrder.RowHeight(-1) = gROWHEIGHT
    
            '## DB에 결과저장
            If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                Res = SaveTransData(gRow, spdOrder)
    
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "저장실패", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
    
                          SQL = "Update PATRESULT Set                                                               " & vbCrLf
                    SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                    SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                    SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                    SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                    SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
            End If
        End If
    End With

Exit Sub

Err:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SerialRcvData_RP500" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_SELEXON()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer
    Dim strTemp1        As String
    Dim strTemp2        As String

On Error GoTo RST
    
    With frmInterface
        strRcvBuf = RcvBuffer

        Call SetSQLData("RCV", strRcvBuf, "A")
        
        strSeq = Trim(Mid(strRcvBuf, 1, 4))         'Data Count
        gHOSP.USERID = Trim(Mid(strRcvBuf, 5, 16))  'Operator ID
        strBarno = Trim(Mid(strRcvBuf, 21, 16))     'Patient ID
        'mResult.RsltDate = Trim(Mid(strRcvBuf, 37, 19))     'RsltDate/RsltTime
        strIntBase = Trim(Mid(strRcvBuf, 56, 9))     'Marker  (hs-CRP, SingleRAW, CK-MB)
        strIntResult = Trim(Mid(strRcvBuf, 65, 13))  'Result
        
        '-- 결과정보
        With mResult
            .BarNo = strBarno
            .RsltDate = Format(Now, "yyyy-mm-dd")
            .RsltTime = Format(Now, "hh:mm:ss")
            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
        End With
        
        '-- 결과환자정보
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        If gRow <= 0 Then
            Exit Sub
        End If
        
        strState = "O"
            
        '-- 검사결과처리 프로세스
        If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
            If strState = "" Or strState = "O" Then
                strState = ""
            End If
            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                strState = "R"
            Else
                If strState = "" Then
                    strState = ""
                End If
            End If
        End If
        
        spdOrder.RowHeight(-1) = gROWHEIGHT

        '(%WBC% * %NEUT%) / 100
        Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
    
        '## DB에 결과저장
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update UB_PATRESULT Set    " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'      " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'" & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
            
            spdOrder.Row = gRow
            spdOrder.Col = colCHECKBOX
            spdOrder.Value = 0
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_SELEXON" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_HITACHI7020()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strTC           As String
    Dim strTG           As String
    Dim strHDL          As String
    Dim strBUN          As String
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strBCRatio      As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                '## ANY 수신
                Case ">", "?", "@"
                    Call SendData(SndMore)
                '## Result
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Call SendData(SndMore)
                
                '## TS inquiry(오더 요청)
                Case ";"
                    ';N     1   1                            
                    strFunc = Mid(strRcvBuf, 2, 1)              ' Function

                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                         Exit Sub
                    End If

                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid$(strRcvBuf, 9, 1)
                    strTubePos = Mid$(strRcvBuf, 10, 3)
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    If gHOSP.BARUSE = "Y" Then
                        '바코드 사용
                        sFunc = Mid(strRcvBuf, 2, 40)
                    Else
                        '바코드 미사용 (바코드 13자리를 '#'으로 변경함)
                        sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    End If
                    
                    With mOrder
                        .BarNo = strBarno
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_HITACHI7020(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    
                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration 데이터는 무시함
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    '## QC
                    If UCase(strFunc) = "F" Then
                        '-- 장비로 전송
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    mOrder.Seq = Trim(strSeq)
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If
                    
                    For ii = 45 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        strIntBase = Format(strIntBase, "00")
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 5))
                        strIntResult = strResult
                        
                        
                        '검사명으로 채널을 찾아와서 비교한다.
                        If strIntBase = gTC Then
                            strTC = strResult
                        End If
                        If strIntBase = gTG Then
                            strTG = strResult
                        End If
                        If strIntBase = gHDL Then
                            strHDL = strResult
                        End If
                        If strIntBase = gBUN Then
                            strBUN = strResult
                        End If
                        If strIntBase = gCREA Then
                            strCREA = strResult
                        End If
                        
                        '-- 검사마스터 정보 가져오기
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        
                        'RA정량
'                        If strIntBase = "20" Then
'                            'RA정성
'                            strIntBase = "99"
'                            If IsNumeric(strResult) Then
'                                If strResult > 15 Then
'                                    strResult = "Positive"
'                                Else
'                                    strResult = "Negative"
'                                End If
'                            End If
'
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
                        
                    Next
                    
                    'LDL 계산
'''                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
'''                        strIntBase = gLDLC
'''                        strResult = strTC - ((strTG / 5) + strHDL)
'''                        If strResult < 0 Then
'''                            strResult = "0"
'''                        End If
'''                        strIntResult = strResult
'''                        strTC = ""
'''                        strTG = ""
'''                        strHDL = ""
'''
'''                        '-- 검사결과처리 프로세스
'''                        If strIntBase <> "" And strResult <> "" Then
'''                            If strState = "" Or strState = "O" Then
'''                                strState = ""
'''                            End If
'''                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'''                                strState = "R"
'''                            Else
'''                                If strState = "" Then
'''                                    strState = ""
'''                                End If
'''                            End If
'''                        End If
'''                    End If
                    
                    'eGFR계산
'''                    If strCREA <> "" And IsNumeric(strCREA) Then
'''                        '18세 이상만 적용
'''                        If mPatient.AGE <> "" And IsNumeric(mPatient.AGE) Then
'''                            If CCur(mPatient.AGE) > 18 Then
'''                                streGFR = ""
'''
'''                                '-- MDRD 공식
'''                                'If mPatient.SEX = "M" Then
'''                                '    streGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203)
'''                                'ElseIf mPatient.SEX = "F" Then
'''                                '    streGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
'''                                'End If
'''
'''                                '--IDMS-MDRD 공식
'''                                If mPatient.SEX = "M" Then
'''                                    streGFR = 175 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203)        'MDRD 공식
'''                                Else 'If mPatient.SEX = "F" Then
'''                                    streGFR = 175 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
'''                                End If
'''
'''                                If streGFR <> "" Then
'''                                    'If streGFR <= 120 Then
'''                                    '    streGFR = Round(streGFR, 2)
'''                                    'ElseIf streGFR > 120 Then
'''                                    '    streGFR = ">120"
'''                                    'End If
'''
'''                                    strIntBase = geGFR
'''                                    strResult = streGFR
'''                                    strIntResult = strResult
'''                                    'strCREA = ""
'''                                    streGFR = ""
'''
'''                                    '-- 검사결과처리 프로세스
'''                                    If strIntBase <> "" And strResult <> "" Then
'''                                        If strState = "" Or strState = "O" Then
'''                                            strState = ""
'''                                        End If
'''                                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'''                                            strState = "R"
'''                                        Else
'''                                            If strState = "" Then
'''                                                strState = ""
'''                                            End If
'''                                        End If
'''                                    End If
'''                                End If
'''                            End If
'''                        End If
'''                    'Else
'''                    '    strIntBase = geGFR
'''                    '    strEGFR = "Error"
'''                    End If
                    
                    
'''                    'BCRATIO
'''                    If strCREA <> "" And IsNumeric(strCREA) And strBUN <> "" And IsNumeric(strBUN) Then
'''                        strBCRatio = ""
'''                        strBCRatio = strBUN / strCREA
'''
'''                        If strBCRatio <> "" Then
'''                            strIntBase = gBCRatio
'''                            strResult = strBCRatio
'''                            strIntResult = strResult
'''                            strCREA = ""
'''                            strBUN = ""
'''                            strBCRatio = ""
'''
'''                            '-- 검사결과처리 프로세스
'''                            If strIntBase <> "" And strResult <> "" Then
'''                                If strState = "" Or strState = "O" Then
'''                                    strState = ""
'''                                End If
'''                                If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'''                                    strState = "R"
'''                                Else
'''                                    If strState = "" Then
'''                                        strState = ""
'''                                    End If
'''                                End If
'''                            End If
'''                        End If
'''                    End If
                    
                    Call SendData(SndMore)
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_H7020" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ResultProcess(ByVal pBarno As String, ByVal pIntBase As String, ByVal pResult As String, ByVal pIntResult As String, Optional ByVal pFlag As String) As Boolean
    Dim RS_L            As ADODB.Recordset
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strSeqNo        As String   '검사순번
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim strCheck        As String   '검사오더체크
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    'Dim strIntResult    As String   '수신한 결과(정량)
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strRstType      As String
    Dim i               As Integer
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCol          As Integer  '결과컬럼 갯수
    
    Dim strIntResult    As String   '변환된 수치결과
    Dim strChrResult    As String   '변환된 문자결과
    Dim strResult       As String   '최종결과
    
    ResultProcess = False
    
    strSeqNo = ""
    strTestCode = ""
    strTestName = ""
    strAbbrName = ""
    intResPrecUse = -1
    intResPrec = -1
    strAMRResult = ""
    
    strIntResult = ""
    strChrResult = ""
    
    mResult.TestQCCd = ""
    
    SQL = ""
    SQL = SQL & "SELECT TESTNAME,ABBRNAME,EQPMASTER.SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC,RESTYPE   " & vbCrLf
'    SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7               " & vbCrLf
'    SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7              " & vbCrLf
'    SQL = SQL & "     , AMRLimit8,  AMRLimit9,  AMRLimit10,  AMRLimit11,  AMRLimit12,  AMRLimit13,  AMRLimit14          " & vbCrLf
'    SQL = SQL & "     , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14         " & vbCrLf
'    SQL = SQL & "     , AMRINResult                                                                                     " & vbCrLf
    SQL = SQL & ", (SELECT TOP 1 TESTMASTER.TESTCODE "
    SQL = SQL & "     FROM TESTMASTER "
    SQL = SQL & "    WHERE TESTMASTER.RSLTCHANNEL = EQPMASTER.RSLTCHANNEL"
    'If gPatOrdCd <> "" Then
    '    SQL = SQL & "      AND TESTMASTER.TESTCODE in (" & gPatOrdCd & ") " & vbCrLf
    'End If
    SQL = SQL & "  ) AS TESTCODE "
    SQL = SQL & "  FROM EQPMASTER LEFT JOIN AMRMASTER                                                                   " & vbCrLf
    SQL = SQL & "   ON (EQPMASTER.RSLTCHANNEL = AMRMASTER.RSLTCHANNEL)                                                  " & vbCrLf
    SQL = SQL & " WHERE EQPMASTER.EQUIPCD     = '" & gHOSP.MACHCD & "'                                                  " & vbCrLf
    If mResult.RackNo <> "" Then
        SQL = SQL & "   AND EQPMASTER.GUBUN       = '" & mResult.RackNo & "'                                            " & vbCrLf
    End If
    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = '" & pIntBase & "'                                                      " & vbCrLf
    'If gPatOrdCd <> "" Then
    '    SQL = SQL & "   AND EQPMASTER.TESTCODE in (" & gPatOrdCd & ") "
    'End If
    
    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        strSeqNo = Trim(RS_L.Fields("SEQNO"))
        strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
        mResult.TestQCCd = strTestCode
        strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
        strAbbrName = Trim(RS_L.Fields("ABBRNAME")) & ""
        
        '-- 소수점변환 사용여부와 변환자리수
        intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
        If Trim(RS_L.Fields("RESPREC")) = "" Then
            intResPrec = 0
        Else
            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
        End If
        
        '-- 성별로 참고치를 비교하여 판정한다.
'''        If mPatient.SEX = "M" Then
'''            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
'''            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
'''        ElseIf mPatient.SEX = "F" Then
'''            strLow = Trim(RS_L.Fields("REFFLOW")) & ""
'''            strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
'''        Else
'''            '-- 남자참고치를 기본으로 한다
'''            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
'''            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
'''        End If
                
        '사용결과  (0:수치,1:판정,2:수치/판정)
        strResType = Trim(RS_L.Fields("RESTYPE")) & ""
        strIntResult = ""
        
        '-- 검사결과가 수치형일경우
        If strResType = 0 Then
            '--- 로우데이터로 결과판정
'''            strJudge = ""
'''            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
'''                If IsNumeric(pIntResult) Then
'''                    If CCur(pIntResult) > CCur(strLow) And CCur(pIntResult) < CCur(strHigh) Then
'''                        strJudge = ""
'''                    ElseIf CCur(strHigh) <= CCur(pIntResult) Then
'''                        strJudge = "H"
'''                    ElseIf CCur(strLow) >= CCur(pIntResult) Then
'''                        strJudge = "L"
'''                    End If
'''                End If
'''            End If
            
            '-- 소수점 처리
            strMachResult = pIntResult
            If intResPrecUse = 1 Then
                For i = 0 To intResPrec
                    If i = 0 Then
                        strResType = "#0"
                    ElseIf i = 1 Then
                        strResType = strResType & ".0"
                    Else
                        strResType = strResType & "0"
                    End If
                Next
                strIntResult = Format(pIntResult, strResType)
            Else
                strIntResult = pIntResult
            End If
            
            '-- 로우데이터로 AMR 적용 (수치형)
'''            If IsNumeric(pIntResult) Then
'''                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
'''                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
'''                        strIntResult = Trim(RS_L.Fields("AMRRESULT1"))
'''                    End If
'''                End If
'''                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
'''                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
'''                        strIntResult = Trim(RS_L.Fields("AMRRESULT2"))
'''                    End If
'''                End If
'''                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
'''                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
'''                        strIntResult = Trim(RS_L.Fields("AMRRESULT3"))
'''                    End If
'''                End If
'''                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
'''                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
'''                        strIntResult = Trim(RS_L.Fields("AMRRESULT4"))
'''                    End If
'''                End If
'''                If strIntResult = "" Then
'''                    strIntResult = pIntResult
'''                End If
'''            Else
'''                strIntResult = pIntResult
'''            End If
            
            If strIntResult <> "" Then
                strResult = strIntResult
            Else
                strResult = pIntResult
            End If

            If strResult = "" Then
                strResult = pResult
            End If
            
'        '-- 검사결과가 문자형일경우
'        ElseIf strResType = 1 Then
'            If pResult <> "" Then
'                '-- AMR 적용 (문자형 단문)
'                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT5"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT6"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT7"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT8"))
'                    End If
'                End If
'
'                '-- AMR 적용 (문자형 장문)
'                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT9"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT10"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT11"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT12"))
'                    End If
'                End If
'
'                If strChrResult = "" Then
'                    strChrResult = pResult
'                End If
'            Else
'                strChrResult = pResult
'            End If
'
'            If strChrResult <> "" Then
'                strResult = strChrResult
'            Else
'                strResult = pResult
'            End If
'
'            If strResult = "" Then
'                strResult = pIntResult
'            End If
'
'        '-- 검사결과가 수치+문자형일경우
'        ElseIf strResType = 2 Then
'            '-- 소수점 처리
'            strMachResult = pIntResult
'            If intResPrecUse = 1 Then
'                For i = 0 To intResPrec
'                    If i = 0 Then
'                        strResType = "#0"
'                    ElseIf i = 1 Then
'                        strResType = strResType & ".0"
'                    Else
'                        strResType = strResType & "0"
'                    End If
'                Next
'                pIntResult = Format(pIntResult, strResType)
'            End If
'
'            '-- AMR 적용 (수치형)
'            If IsNumeric(pIntResult) Then
'                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
'                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
'                        strIntResult = Trim(RS_L.Fields("AMRRESULT1"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
'                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
'                        strIntResult = Trim(RS_L.Fields("AMRRESULT2"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
'                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
'                        strIntResult = Trim(RS_L.Fields("AMRRESULT3"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
'                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
'                        strIntResult = Trim(RS_L.Fields("AMRRESULT4"))
'                    End If
'                End If
'
'                If strIntResult = "" Then
'                    strIntResult = pIntResult
'                End If
'            Else
'                strIntResult = pIntResult
'            End If
'
'            '-- AMR 적용 (문자형)
'            If pResult <> "" Then
'                '-- AMR 적용 (문자형 단문)
'                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT5"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT6"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT7"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT8"))
'                    End If
'                End If
'
'                '-- AMR 적용 (문자형 장문)
'                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT9"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT10"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT11"))
'                    End If
'                End If
'                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
'                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
'                        strChrResult = Trim(RS_L.Fields("AMRRESULT12"))
'                    End If
'                End If
'
'                If strChrResult = "" Then
'                    strChrResult = pResult
'                End If
'
'            Else
'                strChrResult = pResult
'            End If
'
'            '수치결과 포함
'            '0:사용안함, 1:정성(정량), 2:정량(정성)
'            If strIntResult <> "" And strChrResult <> "" Then
'                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
'                    strResult = strChrResult & "(" & strIntResult & ")"
'                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
'                    strResult = strIntResult & "(" & strChrResult & ")"
'                End If
'            Else
'                If strChrResult <> "" Then
'                    strResult = strChrResult
'                ElseIf strIntResult <> "" Then
'                    strResult = strIntResult
'                End If
'            End If
        End If
        
        With frmInterface
            '-- 결과Row 추가
            intRstRow = .spdResult.DataRowCnt + 1
            If .spdResult.MaxRows < intRstRow Then
                .spdResult.MaxRows = intRstRow
            End If
    
            '-- 진행상태 표시("결과")
            SetText .spdOrder, "장비결과", gRow, colSTATE
    
            '-- 메인화면 결과값 표시
''            For intCol = colSTATE + 1 To .spdOrder.MaxCols
''                '약어로 폼로드시 갖고있던 약어값과 비교한다.
''                If strAbbrName = gArrEQP(intCol - colSTATE, 6) Then
''                    SetText .spdOrder, strResult, gRow, intCol
''
''                    '-- H/L 색깔표시
''                    If strJudge = "H" Then
''                        .spdOrder.Row = gRow
''                        .spdOrder.Col = intCol
''                        .spdOrder.ForeColor = vbRed
''                    ElseIf strJudge = "L" Then
''                        .spdOrder.Row = gRow
''                        .spdOrder.Col = intCol
''                        .spdOrder.ForeColor = vbBlue
''                    Else
''                        .spdOrder.Row = gRow
''                        .spdOrder.Col = intCol
''                        .spdOrder.ForeColor = vbBlack
''                    End If
''
''                    '-- EMR에 따라 처방코드와 검사서브코드를 담는다. (코드가 있던 없던간에..)
''                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
''                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
''
''                    Exit For
''                End If
''            Next
    
            '-- 결과 List
            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               '체크
            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '순번
            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
            SetText .spdResult, pIntBase, intRstRow, colRCHANNEL              '장비채널
            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS결과
            '알러지용
            SetText .spdResult, pFlag, intRstRow, colRFLAG                      'FLAG
            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
            
            '-- 이전결과 조회
            strPrevRslt = GetPrevResult(mResult.BarNo, pIntBase, strTestCode)
            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '이전결과
            
            '-- H/L 색깔표시
'''            If strJudge = "H" Then
'''                .spdResult.Row = intRstRow
'''                .spdResult.Col = colRLISRESULT
'''                .spdResult.ForeColor = vbRed
'''                .spdResult.FontBold = True
'''            ElseIf strJudge = "L" Then
'''                .spdResult.Row = intRstRow
'''                .spdResult.Col = colRLISRESULT
'''                .spdResult.ForeColor = vbBlue
'''                .spdResult.FontBold = True
'''            Else
'''                .spdResult.Row = intRstRow
'''                .spdResult.Col = colRLISRESULT
'''                .spdResult.ForeColor = vbBlack
'''                .spdResult.FontBold = False
'''            End If
            
            '-- 결과Count
            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                SetText .spdOrder, "1", gRow, colRCNT
            Else
                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
            End If
        End With
        
        spdOrder.EnhanceStaticCells = False
        
        '-- 로컬 저장
        Call SetLocalDB(gRow, intRstRow, "1", "")
        
        ResultProcess = True
    
    End If
    
End Function

Private Function ResultProcess_Allergy(ByVal pBarno As String, ByVal pIntName As String, ByVal pIntBase As String, ByVal pResult As String, ByVal pIntResult As String, Optional ByVal pFlag As String) As Boolean
    Dim RS_L            As ADODB.Recordset
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strSeqNo        As String   '검사순번
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim strCheck        As String   '검사오더체크
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    'Dim strIntResult    As String   '수신한 결과(정량)
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strRstType      As String
    Dim i               As Integer
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCol          As Integer  '결과컬럼 갯수
    
    Dim strIntResult    As String   '변환된 수치결과
    Dim strChrResult    As String   '변환된 문자결과
    Dim strResult       As String   '최종결과
    
    ResultProcess_Allergy = False
    
    strSeqNo = ""
    strTestCode = ""
    strTestName = ""
    strAbbrName = ""
    intResPrecUse = -1
    intResPrec = -1
    strAMRResult = ""
    
    strIntResult = ""
    strChrResult = ""
    
    mResult.TestQCCd = ""
    
    strOrderCode = gB4C.ORDCODE 'D7460006
    strTestCode = pIntBase
    mResult.TestQCCd = strTestCode
    strTestName = pIntName
    strAbbrName = pIntName
    strResult = pResult
        
    With frmInterface
        '-- 결과Row 추가
        intRstRow = .spdResult.DataRowCnt + 1
        If .spdResult.MaxRows < intRstRow Then
            .spdResult.MaxRows = intRstRow
        End If
        strSeqNo = intRstRow

        '-- 진행상태 표시("결과")
        SetText .spdOrder, "장비결과", gRow, colSTATE

        '-- 메인화면 결과값 표시
        If strTestName = "96M" Then
            SetText .spdOrder, "별지참조", gRow, colITEMS
        End If
        
        strSeqNo = intRstRow - 1
        
        '-- 결과 List
        SetText .spdResult, "1", intRstRow, colRCHECKBOX               '체크
        SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '순번
        SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            '처방코드
        SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '검사코드
        SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '검사코드SUB
        SetText .spdResult, strTestName, intRstRow, colRTESTNM              '검사명
        SetText .spdResult, pIntBase, intRstRow, colRCHANNEL              '장비채널
        SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '장비결과
        SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS결과
        '알러지용
        SetText .spdResult, pFlag, intRstRow, colRFLAG                      'FLAG
        SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '판정
        SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '참고치
    
    End With
    
    spdOrder.EnhanceStaticCells = False
    
    '-- 로컬 저장
    Call SetLocalDB(gRow, intRstRow, "1", "")
    
    ResultProcess_Allergy = True

End Function



'Private Sub SerialRcvData_MEDONIC()
'    '장비 수신 변수
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정성)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'
'    '마스터 변수
'
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'    Dim strQCTemp       As String
'    Dim strRData()      As String
'
'    Dim i               As Integer
'    Dim J               As Integer
'    Dim k               As Integer
'    Dim m               As Integer
'    Dim ii              As Integer
'    Dim intTestNmCnt    As Integer
'    Dim intTestCdCnt    As Integer
'    Dim intOrdCnt       As Integer
'    Dim blnSame         As Boolean
'
'    Dim strTemp1        As String
'    Dim strTemp2        As String
'
'    '계산식 관련
'    Dim strCREA         As String
'    Dim streGFR         As String
'    Dim strFunction     As String
'    Dim strFunc         As String
'    Dim sFunc           As String
'
'    Dim strWBC          As String
'    Dim strNeut         As String
'    Dim strCalChannel   As String
'    Dim strCalCulate    As String
'    Dim varCalCulate    As Variant
'    Dim strCalNm(10)    As String
'    Dim strCalCon(10)   As String
'
'On Error GoTo RST
'
'    strRecvData = Split(RcvBuffer, vbLf)
'    strState = ""
'
'    With frmMain
'        For intCnt = 0 To UBound(strRecvData) - 1
'            strRcvBuf = strRecvData(intCnt)
'
'            Call SetSQLData("RCV", strRcvBuf, "A")
'
'            If InStr(strRcvBuf, "<smpinfo>") > 0 Then
'                strState = "O"
'            End If
'
'            If strState = "O" Then
'                '<p><n>ID</n></p>
'                If InStr(strRcvBuf, "<n>ID</n>") Then
'                    If InStr(strRcvBuf, "<v>") Then
'                        strBarno = mGetP(strRcvBuf, 5, "<")
'                        strBarno = mGetP(strBarno, 2, ">")
'                    Else
'                        strBarno = Format(Now, "yymmddhhmmss")
'                    End If
'                End If
'
'                '<p><n>SEQ</n><v>3060</v></p>
'                If InStr(strRcvBuf, "<n>SEQ</n>") Then
'                    If InStr(strRcvBuf, "<v>") Then
'                        strSeq = mGetP(strRcvBuf, 5, "<")
'                        strSeq = mGetP(strSeq, 2, ">")
'                    Else
'                        strSeq = ""
'                    End If
'
'                    '-- 결과정보
'                    With mResult
'                        .BarNo = strBarno
'                        .RsltDate = Format(Now, "yyyy-mm-dd")
'                        .RsltTime = Format(Now, "hh:mm:ss")
'                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'                    End With
'
'
'                    '-- 결과환자정보
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                    If gRow <= 0 Then
'                        Exit Sub
'                    End If
'                End If
'
'                If strState = "O" Or strState = "R" Then
'                    If InStr(strRcvBuf, "<smpresults>") > 0 Then
'                        strState = "R"
'                    End If
'
'                    If strState = "R" And InStr(strRcvBuf, "<smpresults>") <= 0 Then
'                        '<p><n>RBC</n><v>4.27</v><l>3.50</l><h>5.50</h></p>
'                        strIntBase = mGetP(strRcvBuf, 3, "<")
'                        strIntBase = mGetP(strIntBase, 2, ">")
'
'                        strResult = mGetP(strRcvBuf, 5, "<")
'                        strResult = mGetP(strResult, 2, ">")
'
'
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If strState = "" Or strState = "O" Then
'                                strState = ""
'                            End If
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'                                strState = "R"
'                            Else
'                                If strState = "" Then
'                                    strState = ""
'                                End If
'                            End If
'                        End If
'
'                        RowHeight(-1) = growheight
'                    End If
'                End If
'
'                If InStr(strRcvBuf, "</smpresults>") > 0 Then
'                    strState = "L"
'                End If
'
'
'                '## DB에 결과저장
'                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "L" Then
'                    Res = SaveTransData(gRow, spdOrder)
'
'                    If Res = -1 Then
'                        '-- 저장 실패
'                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                        SetText .spdOrder, "저장실패", gRow, colSTATE
'                    Else
'                        '-- 저장 성공
'                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                        SetText .spdOrder, "저장완료", gRow, colSTATE
'                        SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
'                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
'                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
'                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
'                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
'                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                        If DBExec(AdoCn_Local, SQL) Then
'                            '-- 성공
'                        End If
'                    End If
'                    strState = ""
'                End If
'            End If
'        Next
'    End With
'
'Exit Sub
'
'RST:
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_MEDONIC" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
'
'
'End Sub

Private Sub CalculateTest(ByVal pBarno As String, ByVal pRow As Long, ByVal SPD As Object)
    Dim i, j, k         As Integer
    Dim m               As Integer
    Dim strCalChannel   As String
    Dim strCalCulate    As String
    Dim varCalCulate    As Variant
    Dim strCalNm()      As String
    Dim strCalCon()     As String
    
    Dim strIntBase      As String
    Dim strIntResult    As String
    Dim strResult       As String
    
    m = 0
    ReDim Preserve strCalNm(m)
    ReDim Preserve strCalCon(m)
    
    With SPD
        For i = colSTATE + 1 To spdOrder.MaxCols
            .Row = 0
            .Col = i
            If .FontBold = True Then                                '-- 계산항목은 굵은글자로 표시됨
                strCalChannel = GetChannel(.Text)                       '-- 계산항목의 채널찾기
                strCalCulate = GetCalContents(strCalChannel, "")        '-- 채널로 계산공식 찾아오기 ex) (%WBC% * %NEUT%) / 100
                varCalCulate = Split(strCalCulate, "%")                 '-- "%" 수만큼 배열을 만든다 ex) 검사 한개당 두개의 "%" 가 있다.
                For j = 1 To UBound(varCalCulate) Step 2                '-- 한개의 검사에 "%"가 2개 이므로.. Step 2
                    For k = colSTATE + 1 To .MaxCols
                        .Row = 0
                        .Col = k
                        If .Text = varCalCulate(j) Then                 '-- 검사 약어가 같으면
                            strCalNm(m) = varCalCulate(j)                   '-- 검사 약어 저장
                            strCalCon(m) = GetText(SPD, pRow, k)            '-- 검사 결과 저장
                            '-- 검사결과값
                            If strCalCon(m) = "" Then                       '-- 대상결과값이 하나라도 없으면 연산하지 않는다.
                                m = 0
                                GoTo NextCal
                            Else                                            '-- 대상항목을 검사결과값으로 바꾼다
                                strCalCulate = Replace(strCalCulate, strCalNm(m), strCalCon(m))
                                m = m + 1
                                ReDim Preserve strCalNm(m)
                                ReDim Preserve strCalCon(m)
                            End If
                            Exit For
                        End If
                    Next
                Next
NextCal:
                If m > 0 Then
                    strCalCulate = Replace(strCalCulate, "%", "")
                    
                    '** 계산식으로 연산하여 결과값을 가져온다
                    strIntResult = CFCompute(strCalCulate)
                    '** 계산식으로 연산하여 결과값을 가져온다
                    
                    strIntBase = strCalChannel
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(pBarno, strIntBase, strResult, strIntResult, "") = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    
                End If
            End If
        Next
    End With
    
End Sub

Private Sub SerialRcvData_XN1000()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    Dim strCalChannel   As String
    Dim strCalCulate    As String
    Dim varCalCulate    As Variant
    Dim strCalNm(10)    As String
    Dim strCalCon(10)   As String
    
On Error GoTo RST

    With frmInterface
        strRcvBuf = RcvBuffer

        Call SetSQLData("RCV", strRcvBuf, "A")

        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
                strState = "H"
                
                strWBC = ""
                strNeut = ""
            
            Case "Q"    '## Request Information
                '2Q|1|15^8^            1000001207^B||||20190904144851||||||N
                
                strState = "Q"
                strQState = "Q"
                
                strTemp1 = mGetP(strRcvBuf, 3, "|")
                
                strRackNo = mGetP(strTemp1, 1, "^")
                strTubePos = mGetP(strTemp1, 2, "^")
                strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                
                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                End With
                
                Call GetOrder_XN1000(strBarno, gHOSP.RSTTYPE)
                
            
            Case "P"    '## Patient
                strState = "P"
            
            Case "O"
                strState = "O"
                
                strTemp1 = mGetP(strRcvBuf, 4, "|")
                strBarno = Trim(mGetP(strTemp1, 3, "^"))
                strRackNo = mGetP(strTemp1, 1, "^")
                strTubePos = mGetP(strTemp1, 2, "^")

                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
'                    If Mid(UCase(strBarno), 1, 4) = "XBAR" Then
'                        Exit Sub
'                    End If
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
            Case "R"
                strState = "R"
                
                '7R|1|^^^^WBC^1|6.15|10*3/uL||N||F||||20190904083314

                strTemp1 = mGetP(strRcvBuf, 3, "|")
                strIntBase = mGetP(strTemp1, 5, "^")
                strTemp2 = mGetP(strRcvBuf, 4, "|")
                strFlag = mGetP(strRcvBuf, 7, "|")
                
                If InStr(strTemp2, "^") > 0 Then
                    '## 정성결과 저장
                    strResult = mGetP(strTemp2, 2, "^")
                Else
                    '## 정량결과 저장
                    strIntResult = strTemp2
                End If
                
'                    If strIntBase = "WBC" And IsNumeric(strResult) Then
'                        strWBC = strResult * 1000
'                    End If
'
'                    If strIntBase = "NEUT%" And IsNumeric(strResult) Then
'                        strNeut = strResult / 100
'                    End If
                
                
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                spdOrder.RowHeight(-1) = gROWHEIGHT

'                    If strWBC <> "" And strNeut <> "" Then
'                        ''ANC = (wbc * 1000 * neut%) / 100
'                        strIntBase = "ANC"
'                        strResult = (strWBC * strNeut)
'                        strResult = Format(strResult, "##0")
'                        strWBC = ""
'                        strNeut = ""
'                        GoTo RST
'                    End If
                
            Case "L"
                '-- 계산항목 연산
                Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
                
'''                m = 0
'''                For i = colSTATE + 1 To spdOrder.MaxCols
'''                    spdOrder.Row = 0
'''                    spdOrder.Col = i
'''                    If spdOrder.FontBold = True Then
'''                        '계산항목 채널찾기
'''                        strCalChannel = GetChannel(spdOrder.Text)
'''
'''                        strCalCulate = GetCalContents(strCalChannel, "")
'''                        varCalCulate = Split(strCalCulate, "%")
'''
'''                        For j = 0 To UBound(varCalCulate)
'''                            For k = colSTATE + 1 To spdOrder.MaxCols
'''                                spdOrder.Row = 0
'''                                spdOrder.Col = k
'''                                If spdOrder.Text = varCalCulate(j) Then
'''                                    strCalNm(m) = varCalCulate(j)
'''                                    strCalCon(m) = GetText(spdOrder, gRow, k)
'''
'''                                    strCalCulate = Replace(strCalCulate, strCalNm(m), strCalCon(m))
'''                                    m = m + 1
'''                                    Exit For
'''                                End If
'''                            Next
'''                        Next
'''
'''                        If m > 0 Then
'''                            strCalCulate = Replace(strCalCulate, "%", "")
'''                            strIntResult = mCalP(strCalCulate)
'''
''''
'''                            strIntBase = strCalChannel
'''                            '-- 검사결과처리 프로세스
'''                            If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
'''                                If strState = "" Or strState = "O" Then
'''                                    strState = ""
'''                                End If
'''                                If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
'''                                    strState = "R"
'''                                Else
'''                                    If strState = "" Then
'''                                        strState = ""
'''                                    End If
'''                                End If
'''                            End If
'''
'''
'''                        End If
'''                    End If
'''                Next
                
                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
        End Select
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_THUNDERBOLT()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    Dim strCalChannel   As String
    Dim strCalCulate    As String
    Dim varCalCulate    As Variant
    Dim strCalNm(10)    As String
    Dim strCalCon(10)   As String
    
On Error GoTo RST

    With frmInterface
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")
    
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
        
            Select Case strType
                Case "H"    '## Header
                    strState = "H"
                Case "Q"    '## Request Information
                    strState = "Q"
                    strQState = "Q"
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    Call GetOrder_THUNDERBOLT(strBarno, gHOSP.RSTTYPE)
                    mPNo = 1
                    mOCnt = 1
                
            
                Case "P"    '## Patient
                    strState = "P"
            
                Case "O"
                    strState = "O"
                    strBarno = mGetP(strRcvBuf, 3, "|")
    
                    'If strOldBarno <> strBarno Then
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                    '        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    
                        '-- 결과환자정보
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    'End If
                    
                    'strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                Case "R"
                    strState = "R"
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    
                    If Mid(strIntResult, 1, 1) = "-" Then
                        strIntResult = "0.01"
                    End If
                    
                    'If Mid(strIntResult, 1, 1) = "-" Then
                    '    strResult = "Negative(0.01)"
                    'Else
                    '    strResult = strResult & "(" & strIntResult & ")"
                    'End If
                    
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT
                
                Case "L"
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_THUNDERBOLT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_MINIVIDAS()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String
    Dim blnSame         As Boolean

On Error GoTo RST

    With frmInterface
        Call SetSQLData("RCV", RcvBuffer, "A")
        strRData = Split(RcvBuffer, "|")
        If UBound(strRData) = 0 Then
            Exit Sub
        End If
        
        '-- Sample No
        strTmp = Trim(strRData(4))
        If Mid(strTmp, 2, 2) <> "ci" Then
            Exit Sub
        End If
        strBarno = Trim(Mid(strTmp, 4))
        
        '-- 장비채널
        strTmp = Trim(strRData(5))
        If Mid(strTmp, 2, 2) <> "rt" Then
            Exit Sub
        End If
        strIntBase = Trim(Mid(strTmp, 4))
        
        '-- 검사결과(정성)
        strTmp = Trim(strRData(9))
        If Mid(strTmp, 2, 2) <> "ql" Then
            Exit Sub
        End If
        strResult = Trim(Mid(strTmp, 4))
        
        '-- 검사결과(정량)
        strTmp = Trim(strRData(10))
        If Mid(strTmp, 2, 2) <> "qn" Then
            Exit Sub
        End If
        strIntResult = Trim(Mid(strTmp, 4))
                
        '정량결과 Flag
        If LEFT(strIntResult, 1) = ">" Or LEFT(strIntResult, 1) = "<" Then
            strFlag = LEFT(strIntResult, 1)
            strIntResult = Trim(Mid(strIntResult, 2))
        End If
        
'        If InStr(sTmp, " ") > 0 Then
'            tmpData() = Split(sTmp, " ")
'
'            If UBound(tmpData) > 0 Then
'                sQN = tmpData(0)
'            End If
'
'            sQN = sFlag & sQN
'
'            If Trim(SQL) <> "" And Trim(sQN) <> "" Then
'                sRst = SQL & "(" & sQN & ")"
'            ElseIf sQN <> "" Then
'                sRst = sQN
'            Else
'                sRst = SQL
'            End If
'        Else
'            sQN = sTmp
'
'            If Trim(SQL) <> "" And Trim(sQN) <> "" Then
'                sRst = SQL & "(" & sQN & ")"
'            ElseIf sQN <> "" Then
'                sRst = sQN
'            Else
'                sRst = SQL
'            End If
'        End If
        
        
        With mResult
            .BarNo = strBarno
            .RsltDate = Format(Now, "yyyy-mm-dd")
            .RsltTime = Format(Now, "hh:mm:ss")
            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
        End With
                    
        '-- 결과환자정보
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        If gRow <= 0 Then
            Exit Sub
        End If
                
        '-- 검사결과처리 프로세스
        If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
            If strState = "" Or strState = "O" Then
                strState = ""
            End If
            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                strState = "R"
            Else
                If strState = "" Then
                    strState = ""
                End If
            End If
        End If
        
        spdOrder.RowHeight(-1) = gROWHEIGHT
    
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
            
            spdOrder.Row = gRow
            spdOrder.Col = colCHECKBOX
            spdOrder.Value = 0
        End If
    
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_THUNDERBOLT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_MULTIPLATE()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "O"
                    strBarno = mGetP(strRcvBuf, 2, "|")

                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(strRcvBuf, 2, "|")
                    strResult = mGetP(strRcvBuf, 3, "|")
                    '단위제거
                    strResult = mGetP(strResult, 1, " ")
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_CA800()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim sBC$, sLC$
    Dim strTemp$
    Dim strInfo$
    
    
On Error GoTo RST

    strRcvBuf = strRecvData(1) 'strBuffer
    
    Call SetSQLData("RCV", strRcvBuf, "A")
    
    sBC = Mid(strRcvBuf, 1, 2)
    sLC = Mid(strRcvBuf, 3, 1)
    
    With frmInterface
        Select Case sBC
            'R2210101U0904191511000501     1000001216B           040      050      [Tx]
            Case "R1", "R2"
                strBarno = Trim(Mid(strRcvBuf, 26, 15))
                
                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                    .RackNo = Trim(Mid(strRcvBuf, 20, 4))
                    .TubePos = Trim(Mid(strRcvBuf, 24, 2))
                End With
                
                Call GetOrder_CA800(strBarno, gHOSP.RSTTYPE)
                
                strState = "Q"
                
            Case "D1"
                strBarno = Trim(Mid(strRcvBuf, 26, 15))
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = Trim(Mid(strRcvBuf, 20, 4))
                    .TubePos = Trim(Mid(strRcvBuf, 24, 2))
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                        
                strTemp = Mid(strRcvBuf, 53)
                
                For i = 1 To Len(strTemp) Step 9
                    strIntBase = Mid$(strTemp, i, 3)
                    strType = Mid$(strIntBase, 3, 1)
                    strResult = Trim$(Mid$(strTemp, i + 3, 5))
                    strFlag = Trim(Mid$(strTemp, i + 8, 1))
                    strInfo = GetInfo(strFlag)
                    
                    Select Case strType
                        Case "1"    '## Time
                            strResult = Trim$(Format$(strResult, "@@@@.@"))
                            If strFlag = "*" Or InStr(strResult, "*") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                        Case "2"    '## Activity percent/concentration
                            strResult = Trim$(Format$(strResult, "@@@.@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            ElseIf Mid$(strIntBase, 1, 2) = "04" Then
                                '   - PT %값이 100이상이면 의미없는 결과라서 "100이상"으로 결과저장
                                '     하는것으로 수정
                                strResult = IIf(Val(strResult) > 100, ">100", strResult)
                            End If
                        Case "3"    '## Ratio
                            strResult = Trim$(Format$(strResult, "@.@@@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                        Case "4"    '## INR
                            strResult = Trim$(Format$(strResult, "@@.@@"))
                            If strFlag = "*" Or strResult = "" Or InStr(strResult, "-") > 0 Then
                                strResult = "" 'IISERROR
                            End If
                    End Select
                    
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                Next
                
                spdOrder.RowHeight(-1) = gROWHEIGHT

                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
        End Select
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Result Flags에 대한 상세설명 조회
'-----------------------------------------------------------------------------'
Private Function GetInfo(ByVal pFlag As String)
    Dim strInfo     As String

    If pFlag = "" Then Exit Function

    Select Case pFlag
        Case "+":   strInfo = "Over the upper control limit"
        Case "-":   strInfo = "Under the lower control limit"
        Case "*":   strInfo = "Analysis error occurred, disparate data of mean data occurred, or Fbg was over analysis range."
        Case "!":   strInfo = "Coagulation time was obtained by re-dilution analysis."
        Case ">":   strInfo = "Over the upper report limit."
        Case "<":   strInfo = "Under the lower report limit."
    End Select

    GetInfo = strInfo
End Function

Private Sub SerialRcvData_CA800_ASTM()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_CA800(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim(mGetP(strTemp1, 3, "^"))

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    
                    strTemp2 = mGetP(strRcvBuf, 5, "|")
                    strFlag = mGetP(strRcvBuf, 8, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strResult = strTemp2
                        strIntResult = strTemp2
                    End If
                    
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XP300()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

   ' ReDim Preserve strRData(UBound(strRecvData))
    
   ' strRData = strRecvData
    
    strRData = Split(RcvBuffer, vbCr)
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = Trim(mGetP(strTemp1, 1, "^"))
                    strTubePos = Trim(mGetP(strTemp1, 2, "^"))
                    strBarno = Trim(mGetP(strRcvBuf, 3, "^"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    strResult = ""
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## 정량결과 저장
                        strIntResult = strTemp2
                    End If
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XP300" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_URINSCAN()
    Dim RS_Q            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim POS             As Integer
    Dim strLABLOT       As String
    Dim strLABAVE       As String
    Dim strLABLOW       As String
    Dim strLABMAX       As String
    
    Dim strDate         As String
    Dim strDTM          As String
    
On Error GoTo RST

    strDate = Format(Now, "yyyymmdd")
    strDTM = Format(Now, "yyyymmddhhmm")
    
    With frmInterface
        POS = InStr(RcvBuffer, "ID_NO")
        If POS > 0 Then
            RcvBuffer = Replace(RcvBuffer, vbLf, "")
            strRecvData = Split(RcvBuffer, vbCr)
            
            '-- 바코드 번호 찾기
            '10종
            If UBound(strRecvData) >= 16 Then
                strRcvBuf = strRecvData(16)
                strBarno = Mid(strRcvBuf, 4, 13)
            '7종
            Else
                strRcvBuf = strRecvData(13)
                strBarno = Mid(strRcvBuf, 4, 13)
            End If
            
            '-- QC 번호 찾기
            strRcvBuf = strRecvData(1)
            strBarno = mGetP(strRcvBuf, 2, ":")
            
            '   0004-q1, 0005-q2
            mResult.Kind = ""
            If UCase(Mid(mGetP(strBarno, 2, "-"), 1, 1)) = "Q" Then
                mResult.Kind = "QC"
                mResult.LabNab = Mid(mGetP(strBarno, 2, "-"), 2, 1)
            End If
            
            '-- SEQ 번호 찾기
'''            strRcvBuf = strRecvData(2)
'''            strRcvBuf = mGetP(strRcvBuf, 2, ":")
'''            strRcvBuf = mGetP(strRcvBuf, 1, "-")
'''            strSeq = Trim(strRcvBuf)
'''            strBarno = strSeq
            
            '-- SEQ 번호 찾기
            'ID_NO:0013-
            strRcvBuf = strRecvData(1)
            strSeq = Trim(mGetP(strRcvBuf, 2, ":"))
            strSeq = Trim(mGetP(strSeq, 1, "-"))
            strSeq = Val(strSeq)
            'mOrder.Seq = strSeq
            
            
            With mResult
                .BarNo = strBarno
                .Seq = strSeq
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                .PatNo = "QC"
            End With
                    
            If mResult.Kind = "QC" Then
                Call SetPatInfoQC(strBarno, gHOSP.RSTTYPE)
            Else
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            End If
            
            For intCnt = 4 To UBound(strRecvData)
                strRcvBuf = strRecvData(intCnt)
                
                strType = Trim(Mid$(strRcvBuf, 1, 3))
                strIntBase = strType
                strResult = ""

                Select Case strType
                    Case "p.H", "pH", "S.G", "SG", "COL" '## 소숫점 포함 3자리
                        strResult = Trim$(Mid$(strRcvBuf, 4))
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    
                    Case Else
                        strResult = Trim$(Mid$(strRcvBuf, 4, 7))
                        'strResult = Trim(Mid(strRcvBuf, 12))  '-- 정량
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                        
                End Select
                
                strIntResult = strResult
                        
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        If mResult.Kind = "QC" Then
                            '-- Lot번호,평균값,LABLOW,LABMAX
                            strLABLOT = ""
                            strLABAVE = ""
                            strLABLOW = ""
                            strLABMAX = ""
                            
                            SQL = ""
                            SQL = SQL & "SELECT LABLOT, LABAVE, LABLOW, LABMAX                              " & vbCrLf
                            SQL = SQL & "  FROM LABQCMST                                                    " & vbCrLf
                            SQL = SQL & " WHERE LABQCCOD   = '" & gHOSP.MACHCD & "'                         " & vbCrLf
                            SQL = SQL & "   And LABNAB     = '" & mResult.LabNab & "'                       " & vbCrLf
                            SQL = SQL & "   And LABCOD     = '" & mResult.TestQCCd & "'                     " & vbCrLf
                            SQL = SQL & "   And LABADPDTE <= '" & strDate & "'                              " & vbCrLf
                            SQL = SQL & "   And LABADPDTE  = (Select max(LABADPDTE) from LABQCMST           " & vbCrLf
                            SQL = SQL & "                      Where LABQCCOD   = '" & gHOSP.MACHCD & "'    " & vbCrLf
                            SQL = SQL & "                        And LABNAB     = '" & mResult.LabNab & "'  " & vbCrLf
                            SQL = SQL & "                        And LABCOD     = '" & mResult.TestQCCd & "'" & vbCrLf
                            SQL = SQL & "                        And LABADPDTE <= '" & strDate & "')        " & vbCrLf
        
                            Set RS_Q = AdoCn.Execute(SQL, , 1)
                            
                            If Not RS_Q.EOF = True And Not RS_Q.BOF = True Then
                                strLABLOT = Trim(RS_Q.Fields("LABLOT") & "")
                                strLABAVE = Trim(RS_Q.Fields("LABAVE") & "")
                                strLABLOW = Trim(RS_Q.Fields("LABLOW") & "")
                                strLABMAX = Trim(RS_Q.Fields("LABMAX") & "")
                            End If
                            RS_Q.Close
                        
                            '-- 이미 결과가 있는지 확인
                            SQL = ""
                            SQL = SQL & "SELECT LABADPDTE                               " & vbCrLf
                            SQL = SQL & "  FROM LABQCINF                                " & vbCrLf
                            SQL = SQL & " WHERE LABADPDTE = '" & strDate & "'           " & vbCrLf      '조회일자
                            SQL = SQL & "   And LABQCCOD  = '" & gHOSP.MACHCD & "'      " & vbCrLf      '장비명
                            SQL = SQL & "   And LABNAB    = '" & mResult.LabNab & "'    " & vbCrLf      '1,2:QC 레벨
                            SQL = SQL & "   And LABCOD    = '" & mResult.TestQCCd & "'  " & vbCrLf      '검사코드
                            Set RS_Q = AdoCn.Execute(SQL, , 1)
                            
                            If Not RS_Q.EOF = True And Not RS_Q.BOF = True Then
                                'UPDATE
                                SQL = ""
                                SQL = SQL & "Update LABQCINF                                " & vbCrLf
                                SQL = SQL & "   Set LABMZH      = '" & strResult & "'       " & vbCrLf    '결과
                                SQL = SQL & "     , LABDTM      = '" & strDTM & "'          " & vbCrLf    '시간
                                SQL = SQL & "     , LABAVE      = '" & strLABAVE & "'       " & vbCrLf    'AVE
                                SQL = SQL & "     , LABLOW      = '" & strLABLOW & "'       " & vbCrLf    'LOW
                                SQL = SQL & "     , LABMAX      = '" & strLABMAX & "'       " & vbCrLf    'MAX
                                SQL = SQL & " Where LABADPDTE   = '" & strDate & "'         " & vbCrLf    '조회일자
                                SQL = SQL & "   And LABQCCOD    = '" & gHOSP.MACHCD & "'    " & vbCrLf    '장비명
                                SQL = SQL & "   And LABNAB      = '" & mResult.LabNab & "'  " & vbCrLf    '1,2:QC1,QC2"
                                SQL = SQL & "   And LABCOD      = '" & mResult.TestQCCd & "'" & vbCrLf     '검사코드
                                'SQL = SQL & "   And LABLOT     = '" & strLABLOT & "'"        'LOT
                            Else
                                'INSERT
                                SQL = ""
                                SQL = SQL & "Insert Into LABQCINF ("
                                SQL = SQL & "  LABADPDTE"
                                SQL = SQL & " , LABQCCOD"
                                SQL = SQL & " , LABNAB"
                                SQL = SQL & " , LABCOD"
                                SQL = SQL & " , LABMZH"
                                SQL = SQL & " , LABUID"
                                SQL = SQL & " , LABDTM"
                                SQL = SQL & " , LABLOT"
                                SQL = SQL & " , LABAVE"
                                SQL = SQL & " , LABLOW"
                                SQL = SQL & " , LABMAX"
                                SQL = SQL & ")" & vbCrLf
                                SQL = SQL & " Values ("
                                SQL = SQL & "   '" & strDate & "'"
                                SQL = SQL & " , '" & gHOSP.MACHCD & "'"
                                SQL = SQL & " , '" & mResult.LabNab & "'"
                                SQL = SQL & " , '" & mResult.TestQCCd & "'"
                                SQL = SQL & " , '" & strResult & "'"
                                SQL = SQL & " , '" & gHOSP.USERID & "'"
                                SQL = SQL & " , '" & strDTM & "'"
                                SQL = SQL & " , '" & strLABLOT & "'"
                                SQL = SQL & " , '" & strLABAVE & "'"
                                SQL = SQL & " , '" & strLABLOW & "'"
                                SQL = SQL & " , '" & strLABMAX & "'"
                                SQL = SQL & ")" & vbCrLf
                            End If
                            
                            Call SetSQLData("결과저장", SQL, "A")
                            AdoCn.Execute SQL
                            
                            strState = ""
                        Else
                            strState = "R"
                        End If
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                spdOrder.RowHeight(-1) = gROWHEIGHT
            Next
                
            '## DB에 결과저장
            If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                Res = SaveTransData(gRow, spdOrder)

                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "저장실패", gRow, colSTATE
                Else
                    '-- 저장 성공
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "저장완료", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX

                          SQL = "Update PATRESULT Set                                                               " & vbCrLf
                    SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                    SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                    SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                    SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                    SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                strState = ""
                
                spdOrder.Row = gRow
                spdOrder.Col = colCHECKBOX
                spdOrder.Value = 0
            End If
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_URINSCAN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_AVL9180()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim POS             As Integer
        
On Error GoTo RST
    
    With frmInterface
        strRcvBuf = RcvBuffer
        RcvBuffer = ""
        
        Call SetSQLData("RCV", strRcvBuf, "A")
        
        If InStr(strRcvBuf, "Sample: QC") > 0 Then
            strOldBarno = Trim(mGetP(strRcvBuf, 2, ":"))
        End If
        
        'Sample
        'strBarno
        
        If InStr(strRcvBuf, "Na=") > 0 Or InStr(strRcvBuf, "K =") > 0 Or InStr(strRcvBuf, "Cl=") > 0 Then
            
            strIntBase = Trim(Mid(strRcvBuf, 1, 2))
            strResult = Trim(Mid(strRcvBuf, 4, 5))
            
            If strIntBase = "Na" Then
                mResult.BarNo = strOldBarno
                mResult.strNa = strResult
                mResult.strK = ""
                mResult.strCl = ""
            ElseIf strIntBase = "K" Then
                mResult.strK = strResult
                mResult.strCl = ""
            ElseIf strIntBase = "Cl" Then
                mResult.strCl = strResult
            End If
                
            If mResult.strCl <> "" Then
                strOldBarno = ""
                With mResult
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
        
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                                
                For i = 1 To 3
                    strIntBase = ""
                    strResult = ""
                    Select Case i
                        Case 1: strIntBase = "Na": strResult = mResult.strNa
                        Case 2: strIntBase = "K":  strResult = mResult.strK
                        Case 3: strIntBase = "Cl": strResult = mResult.strCl
                    End Select
                    
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                
                Next
                       
                mResult.strNa = ""
                mResult.strK = ""
                mResult.strCl = ""
                
                .spdResult.RowHeight(-1) = gROWHEIGHT ' = 14
                
                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)
    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
    
                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            End If
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AVL9180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_PATHFAST()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim(mGetP(strTemp1, 1, "^"))
                    strSeq = Trim(mGetP(strTemp1, 2, "^"))
                    
                    '-- 결과정보
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        With mResult
                            .BarNo = strBarno
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_PATHFAST" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_VISION()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    strRData = Split(pBuffer, vbLf)
    
    With frmInterface
        For intCnt = 0 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            Call SetSQLData("RCV", strRcvBuf, "A")
            
            If Len(strRcvBuf) > 20 Then
                strIntBase = "ESR"
                strSeq = mGetP(strRcvBuf, 1, vbTab)
                strBarno = mGetP(strRcvBuf, 7, vbTab)
                '-- 18도 사용
                strResult = mGetP(strRcvBuf, 10, vbTab)
                strIntResult = mGetP(strRcvBuf, 10, vbTab)
                'strResult = mGetP(strRcvBuf, 11, vbTab)

                '-- 결과정보
                With mResult
                    .BarNo = strBarno
                    .Seq = strSeq
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                        
                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                spdOrder.RowHeight(-1) = gROWHEIGHT

                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            End If
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_KLITE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ISMART30()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                Case "O"
'                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
'
'                    '-- 결과정보
'                    With mResult
'                        .BarNo = strBarno
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
'                        .RsltDate = Format(Now, "yyyy-mm-dd")
'                        .RsltTime = Format(Now, "hh:mm:ss")
'                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'                    End With
'
'                    '-- 결과환자정보
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                    strState = "O"
'
'                    If gRow <= 0 Then
'                        Exit Sub
'                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_MICROS60()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strTmp = Trim(mGetP(strRcvBuf, 3, "|"))
                    strBarno = Trim(mGetP(strTmp, 1, "^"))
                    strRackNo = Trim(mGetP(strTmp, 2, "^"))
                    strTubePos = Trim(mGetP(strTmp, 3, "^"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "R"
                    strIntBase = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    strResult = Trim$(mGetP(strRcvBuf, 4, "|"))
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_MICROS60" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XL1000I()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    
On Error GoTo RST


'3 2012090230
'sampleNo 5
'patientNo 2012090230
'name
'SEX
'AGE
'sampleType  Whole blood
'department
'hospitalNo
'treatmentRoom
'bedLabel
'diagnosis
'Date 2020 - 12 - 9
'time    14:31:42
'doctor
'Operator admin
'validator
'Memo
'printTime
'outputTime
'PT-S    11.6    s   9.0 13.0
'PT-OD   0.0993063   OD  0   0
'PT-%    104.56  %   80.00   150.00
'PTR 0.95        0.82    1.15
'PT-INR  0.95        0.76    1.15
'APTT    34.6    s   25.0    41.0
'APTTR   1.12        0.77    1.23
'

    With frmInterface
        'RcvBuffer = Replace(RcvBuffer, vbLf, "")
        strRData = Split(RcvBuffer, vbLf)
        
        For intCnt = 0 To UBound(strRData) - 1
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")
        
            If mGetP(strRcvBuf, 1, vbTab) = "sampleNo" Then
'                strBarno = mGetP(strRcvBuf, 2, vbTab)
'                '-- 결과정보
'                With mResult
'                    .BarNo = strBarno
'                    .RackNo = strRackNo
'                    .TubePos = strTubePos
'                    .RsltDate = Format(Now, "yyyy-mm-dd")
'                    .RsltTime = Format(Now, "hh:mm:ss")
'                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'                End With
'
'                '-- 결과환자정보
'                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                strState = "O"
'
'                If gRow <= 0 Then
'                    Exit Sub
'                End If
            
            ElseIf mGetP(strRcvBuf, 1, vbTab) = "patientNo" Then
                strBarno = mGetP(strRcvBuf, 2, vbTab)
                '-- 결과정보
                With mResult
                    .BarNo = strBarno
'                    .RackNo = strRackNo
'                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            ElseIf mGetP(strRcvBuf, 1, vbTab) = "name" Then
            
            Else
                strIntBase = mGetP(strRcvBuf, 1, vbTab)
                strResult = mGetP(strRcvBuf, 2, vbTab)
                strIntResult = strResult
                
                If mGetP(strRcvBuf, 4, vbTab) <> "" Then    '참고치가 비어있지 않으면 검사라고 인식함...
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                End If
                spdOrder.RowHeight(-1) = gROWHEIGHT
            End If
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
            
            spdOrder.Row = gRow
            spdOrder.Col = colCHECKBOX
            spdOrder.Value = 0
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XL1000I" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_LTC52()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    
On Error GoTo RST

    'PT
    '1 4 6 17 10 11 10 33 001214 PT    011.4S 011.0 017.0 00.91R 000.90I 125.2%
    'APTT
    '1 4 6 17 10 11 10 35 000366 APTT  025.2S 023.0 045.0 

    With frmInterface
        strRcvBuf = RcvBuffer
        Call SetSQLData("RCV", strRcvBuf, "A")
        strRData = Split(strRcvBuf, " ")
        
        For i = 1 To UBound(strRData)
            Select Case i
            Case 8
                strBarno = strRData(i)
                
                '-- 결과정보
                With mResult
                    .BarNo = strBarno
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
            Case 9
                strTemp1 = strRData(i)
                
            Case 10 To UBound(strRData) - 1
                If Right(strRData(i), 1) = "S" Or Right(strRData(i), 1) = "R" Or Right(strRData(i), 1) = "%" Then
                    strIntBase = strTemp1 & Right(strRData(i), 1)
                    strResult = Mid(strRData(i), 1, Len(strRData(i)) - 1)
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    spdOrder.RowHeight(-1) = gROWHEIGHT
                End If
            End Select
        Next
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
            
            spdOrder.Row = gRow
            spdOrder.Col = colCHECKBOX
            spdOrder.Value = 0
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_XL1000I" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ARKRAY()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim sQcCheck        As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    strRackNo = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
                    strTubePos = mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^")
                    '일반
                    'O|1|2011200092--------^0005^01|0002|^^^HbA1c|R||||||||||||||00000000^00000000^0^F^------^-------^------^-------^------^-------^------^-------^-------^-------
                    'QC
                    'O|1|------------------^0005^01|0002|^^^HbA1c|C||||||||||||||00000000^00000000^0^F^------^-------^------^-------^------^-------^------^-------^-------^-------
                    strBarno = Replace(strBarno, "-", "")
                    If strBarno = "" Then
                        strBarno = strRackNo & "-" & strTubePos
                    End If
                    sQcCheck = Trim(mGetP(strRcvBuf, 6, "|"))
    
                    If sQcCheck = "C" Then
                        strBarno = "QC" & strBarno
                    Else
                        'If Not IsNumeric(strBarno) Then
                        '    Exit Sub
                        'End If
                    End If
                
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    'R|1|^^^ValueHbA1c|6.7|%||||F|||201610250622
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ARKRAY" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_D10()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        strRcvBuf = RcvBuffer
        Call SetSQLData("RCV", strRcvBuf, "A")
        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"
                strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                strSeq = Trim(mGetP(strRcvBuf, 2, "|"))

                '-- 결과정보
                With mResult
                    .BarNo = strBarno
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                strA1cResult = ""
                strA1cIntBase = ""
                
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
            Case "R"
                strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                strFlag = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                strIntBase = strIntBase & "_" & strFlag
                strIntResult = strResult
                
'                If strIntBase = "A1c_AREA" Then
'                    strA1cIntBase = strIntBase
'                    strA1cResult = strResult
'
'                ElseIf strIntBase = "TOTAL_AREA" Then
'                    If strResult < 1000000 Then
'                        strA1cResult = "*"
'                    End If
'
'                    If strResult > 5000000 Then
'                        strA1cResult = "*"
'                    End If
'
'                        'A1c 저장
'                        strIntBase = strA1cIntBase
'                        strResult = strA1cResult
'                    End If
'                End If

                '-- 검사결과처리 프로세스
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                spdOrder.RowHeight(-1) = gROWHEIGHT

            Case "L"
            
                '## DB에 결과저장
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "저장실패", gRow, colSTATE
                    Else
                        '-- 저장 성공
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "저장완료", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    End If
                    strState = ""
                    
                    spdOrder.Row = gRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
        End Select
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_D10" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_EPOC()
    Dim RS_Q            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strLABLOT       As String
    Dim strLABAVE       As String
    Dim strLABLOW       As String
    Dim strLABMAX       As String
    
    Dim strDate         As String
    Dim strDTM          As String
    
On Error GoTo RST

    strDate = Format(Now, "yyyymmdd")
    strDTM = Format(Now, "yyyymmddhhmm")
    
    strRecvData = Split(RcvBuffer, vbCrLf)
    
    With frmInterface
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")
        
            
            Select Case intCnt
                Case 7
                    If InStr(strRecvData(intCnt), "Patient ID") > 0 Then
                        strBarno = Trim(mGetP(strRecvData(intCnt), 2, ":"))
    
                        '-- 결과정보
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        
                        mResult.LabNab = ""
                        If UCase(Mid(strBarno, 1, 2)) = "QC" Then
                            'QC1....., QC2....., QC3....., QC4.....
                            mResult.PatNo = Mid(strBarno, 1, 3)
                            mResult.LabNab = Mid(strBarno, 3, 1)
                            mResult.Kind = "QC"
                            
                            '-- 결과환자정보
                            Call SetPatInfoQC(strBarno, gHOSP.RSTTYPE)
                        Else
                            mResult.Kind = ""
                            
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        End If
                        
                        strState = "O"
                    
                        If gRow <= 0 Then
                            Exit Sub
                        End If
                    End If
                    
                Case 8 To 70
                    strIntBase = Trim(Mid(strRecvData(intCnt), 1, 9))
                    strResult = Trim(Mid(strRecvData(intCnt), 10, 7))
                    
                    If strIntBase = "Reference" Then
                        Exit For
                    End If
                    
                    If UCase(strIntBase) = "REFERENCE" Then
                        Exit For
                    End If
    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            If mResult.Kind = "QC" Then
                                '-- Lot번호,평균값,LABLOW,LABMAX
                                strLABLOT = ""
                                strLABAVE = ""
                                strLABLOW = ""
                                strLABMAX = ""
                                
                                SQL = ""
                                SQL = SQL & "SELECT LABLOT, LABAVE, LABLOW, LABMAX                              " & vbCrLf
                                SQL = SQL & "  FROM LABQCMST                                                    " & vbCrLf
                                SQL = SQL & " WHERE LABQCCOD   = '" & gHOSP.MACHCD & "'                         " & vbCrLf
                                SQL = SQL & "   And LABNAB     = '" & mResult.LabNab & "'                       " & vbCrLf
                                SQL = SQL & "   And LABCOD     = '" & mResult.TestQCCd & "'                     " & vbCrLf
                                SQL = SQL & "   And LABADPDTE <= '" & strDate & "'                              " & vbCrLf
                                SQL = SQL & "   And LABADPDTE  = (Select max(LABADPDTE) from LABQCMST           " & vbCrLf
                                SQL = SQL & "                      Where LABQCCOD   = '" & gHOSP.MACHCD & "'    " & vbCrLf
                                SQL = SQL & "                        And LABNAB     = '" & mResult.LabNab & "'  " & vbCrLf
                                SQL = SQL & "                        And LABCOD     = '" & mResult.TestQCCd & "'" & vbCrLf
                                SQL = SQL & "                        And LABADPDTE <= '" & strDate & "')        " & vbCrLf
            
                                Set RS_Q = AdoCn.Execute(SQL, , 1)
                                
                                If Not RS_Q.EOF = True And Not RS_Q.BOF = True Then
                                    strLABLOT = Trim(RS_Q.Fields("LABLOT") & "")
                                    strLABAVE = Trim(RS_Q.Fields("LABAVE") & "")
                                    strLABLOW = Trim(RS_Q.Fields("LABLOW") & "")
                                    strLABMAX = Trim(RS_Q.Fields("LABMAX") & "")
                                End If
                                RS_Q.Close
                            
                                '-- 이미 결과가 있는지 확인
                                SQL = ""
                                SQL = SQL & "SELECT LABADPDTE                               " & vbCrLf
                                SQL = SQL & "  FROM LABQCINF                                " & vbCrLf
                                SQL = SQL & " WHERE LABADPDTE = '" & strDate & "'           " & vbCrLf      '조회일자
                                SQL = SQL & "   And LABQCCOD  = '" & gHOSP.MACHCD & "'      " & vbCrLf      '장비명
                                SQL = SQL & "   And LABNAB    = '" & mResult.LabNab & "'    " & vbCrLf      '1,2:QC 레벨
                                SQL = SQL & "   And LABCOD    = '" & mResult.TestQCCd & "'  " & vbCrLf      '검사코드
                                Set RS_Q = AdoCn.Execute(SQL, , 1)
                                
                                If Not RS_Q.EOF = True And Not RS_Q.BOF = True Then
                                    'UPDATE
                                    SQL = ""
                                    SQL = SQL & "Update LABQCINF                                " & vbCrLf
                                    SQL = SQL & "   Set LABMZH      = '" & strResult & "'       " & vbCrLf    '결과
                                    SQL = SQL & "     , LABDTM      = '" & strDTM & "'          " & vbCrLf    '시간
                                    SQL = SQL & "     , LABAVE      = '" & strLABAVE & "'       " & vbCrLf    'AVE
                                    SQL = SQL & "     , LABLOW      = '" & strLABLOW & "'       " & vbCrLf    'LOW
                                    SQL = SQL & "     , LABMAX      = '" & strLABMAX & "'       " & vbCrLf    'MAX
                                    SQL = SQL & " Where LABADPDTE   = '" & strDate & "'         " & vbCrLf    '조회일자
                                    SQL = SQL & "   And LABQCCOD    = '" & gHOSP.MACHCD & "'    " & vbCrLf    '장비명
                                    SQL = SQL & "   And LABNAB      = '" & mResult.LabNab & "'  " & vbCrLf    '1,2:QC1,QC2"
                                    SQL = SQL & "   And LABCOD      = '" & mResult.TestQCCd & "'" & vbCrLf     '검사코드
                                    'SQL = SQL & "   And LABLOT     = '" & strLABLOT & "'"        'LOT
                                Else
                                    'INSERT
                                    SQL = ""
                                    SQL = SQL & "Insert Into LABQCINF ("
                                    SQL = SQL & "  LABADPDTE"
                                    SQL = SQL & " , LABQCCOD"
                                    SQL = SQL & " , LABNAB"
                                    SQL = SQL & " , LABCOD"
                                    SQL = SQL & " , LABMZH"
                                    SQL = SQL & " , LABUID"
                                    SQL = SQL & " , LABDTM"
                                    SQL = SQL & " , LABLOT"
                                    SQL = SQL & " , LABAVE"
                                    SQL = SQL & " , LABLOW"
                                    SQL = SQL & " , LABMAX"
                                    SQL = SQL & ")" & vbCrLf
                                    SQL = SQL & " Values ("
                                    SQL = SQL & "   '" & strDate & "'"
                                    SQL = SQL & " , '" & gHOSP.MACHCD & "'"
                                    SQL = SQL & " , '" & mResult.LabNab & "'"
                                    SQL = SQL & " , '" & mResult.TestQCCd & "'"
                                    SQL = SQL & " , '" & strResult & "'"
                                    SQL = SQL & " , '" & gHOSP.USERID & "'"
                                    SQL = SQL & " , '" & strDTM & "'"
                                    SQL = SQL & " , '" & strLABLOT & "'"
                                    SQL = SQL & " , '" & strLABAVE & "'"
                                    SQL = SQL & " , '" & strLABLOW & "'"
                                    SQL = SQL & " , '" & strLABMAX & "'"
                                    SQL = SQL & ")" & vbCrLf
                                End If
                                
                                Call SetSQLData("결과저장", SQL, "A")
                                AdoCn.Execute SQL
                                
                                strState = ""
                            Else
                                strState = "R"
                            End If
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT
                    
            End Select
        Next
    
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "저장실패", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
            
            spdOrder.Row = gRow
            spdOrder.Col = colCHECKBOX
            spdOrder.Value = 0
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_EPOC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_BIOLYTE()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

   ' ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = Split(RcvBuffer, vbCr)
    'strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 3, 1)

            'strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    If InStr(strResult, "^") > 0 Then
                        '## 정성결과 저장
                        strResult = mGetP(strResult, 2, "^")
                    End If
                    
                    strIntResult = strResult
                    
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_BIOLYTE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_YUMIZEN()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^289645146||ALL||||||||O<CR><ETX>F7<CR><LF
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))

                    With mOrder
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder(Trim$(strBarno), gHOSP.RSTTYPE)
                    'Call GetOrder_YUMIZEN(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))

                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    ''Call SetCommStatus("R", strBarno, lstComStatus)
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    strIntResult = strResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                        
                        spdOrder.Row = gRow
                        spdOrder.Col = colCHECKBOX
                        spdOrder.Value = 0
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_STAGO()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_STAGO(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = mGetP(strTemp1, 1, "^")
                    strSeq = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    strBarno = Replace(strBarno, "_", "1")
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .Seq = strSeq
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    Select Case strFlag
                        Case "F"    '## 정량
                            strIntResult = strIntResult
                        Case "I"    '## 정성
                            Select Case Mid$(strIntResult, 1, 1)
                                Case "N":   strResult = "Negative"
                                Case "G":   strResult = "GRAYZONE"
                                Case "R":   strResult = "Positive"
                                Case "P":   strResult = "Positive"
                            End Select
                    End Select
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_STAGO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ACCESS2()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^190807015||ALL||||||||O

                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_ACCESS2(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "P"    '## Patient
                Case "O"
                    '3O|1|190807015|^1403^1|^^^HCG5^1|||||||||||Serum||||||||||F
                    '4R|1|^^^HCG5^1|>1342.00|mIU/mL|0.00 to 5.00^normal|>|N|F||||20190807153839|511896
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strRackNo = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    strRackNo = Format(strRackNo, "0000")
                    strTubePos = Format(strTubePos, "00")
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                Case "R"
                    '4R|1|^^^hLH^1|17.28|mIU/mL||N||F||||20190731123358|511896
                    
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strIntResult = mGetP(strTemp1, 1, "^")
                    strResult = mGetP(strTemp1, 2, "^")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If strResult = "" Then
                        strResult = strIntResult
                    End If
'                    If strIntBase = "HBsAgV3" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 1 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    'HbsAb
'                    '4R|1|^^^HBAb3^1|0.7|mIU/mL||N||F||||20190415103432|510062
'                    ElseIf strIntBase = "HBAb3" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 10 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    'HCV
'                    '4R|1|^^^HCVPLUS^1|0.10^Non-React.|S/CO||N||F||||20190415103620|510062
'                    ElseIf strIntBase = "HCVPLUS" Then
'                        If IsNumeric(strIntResult) Then
'                            If CCur(strIntResult) < 1 Then
'                                strResult = "Negative(" & strIntResult & ")"
'                            Else
'                                strResult = "Positive(" & strIntResult & ")"
'                            End If
'                        End If
'                    Else
'                        strResult = strIntResult
'                    End If
                    
                    
'                    Select Case strFlag
'                        Case "F"    '## 정량
'                            strResult = strIntResult
'                        Case "I"    '## 정성
'                            Select Case Mid$(strIntResult, 1, 1)
'                                Case "N":   strResult = "Negative"
'                                Case "G":   strResult = "GRAYZONE"
'                                Case "R":   strResult = "Positive"
'                                Case "P":   strResult = "Positive"
'                            End Select
'                    End Select
                        
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                Case "L"
                
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ACCESS2" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub RcvData()
    Dim vBuffers As Variant
    
    Select Case UCase(gHOSP.MACHNM)
        '-- 시리얼
        Case "ACCESS2":         Call SerialRcvData_ACCESS2
        Case "AFIAS6":          Call SerialRcvData_AFIAS6
        Case "ARKRAY":          Call SerialRcvData_ARKRAY
        Case "AU480":           Call SerialRcvData_AU480
        Case "AVL9180":         Call SerialRcvData_AVL9180
        Case "BIOLYTE":         Call SerialRcvData_BIOLYTE
        Case "CA800":           Call SerialRcvData_CA800
        Case "CA800_ASTM":      Call SerialRcvData_CA800_ASTM
        Case "EPOC":            Call SerialRcvData_EPOC
        Case "HITACHI7020":     Call SerialRcvData_HITACHI7020
        Case "HITACHI7180":     Call SerialRcvData_HITACHI7180
        Case "HORIBA":          Call SerialRcvData_HORIBA
        'Case "INDIKO":          Call SerialRcvData_INDIKO
        Case "ISMART30":        Call SerialRcvData_ISMART30
        'Case "KLITE":           Call SerialRcvData_KLITE
        Case "LTC52":           Call SerialRcvData_LTC52
        'Case "MEDONIC":         Call SerialRcvData_MEDONIC
        Case "MICROS60":        Call SerialRcvData_MICROS60
        Case "MINIVIDAS":       Call SerialRcvData_MINIVIDAS
        Case "PATHFAST":        Call SerialRcvData_PATHFAST
        Case "RP500":           Call SerialRcvData_RP500
        Case "SELEXON":         Call SerialRcvData_SELEXON
        Case "STAGO":           Call SerialRcvData_STAGO
        Case "THUNDERBOLT":     Call SerialRcvData_THUNDERBOLT
        Case "URINSCAN":        Call SerialRcvData_URINSCAN
        Case "UROMETER":        Call SerialRcvData_UROMETER720
        Case "UROMETER120":     Call SerialRcvData_UROMETER720
        Case "UROMETER720":     Call SerialRcvData_UROMETER720
        Case "XL1000I":         Call SerialRcvData_XL1000I
        Case "XN1000":          Call SerialRcvData_XN1000
        Case "XP300":           Call SerialRcvData_XP300
        Case "YUMIZEN":         Call SerialRcvData_YUMIZEN           '영인과학 HORIBA YUMIZEN H500
        '-- 소켓
        Case "BS360S":          vBuffers = Split(pBuffer, vbCr)
                                Call Sleep(200)
                                If UBound(vBuffers) > 0 Then
                                    Call TCPRcvData_BS360S
                                End If
        Case "BC5180":          Call TCPRcvData_BC5180
        Case "BC6200":          Call TCPRcvData_BC6200
        Case "GENEXPERT":       Call TCPRcvData_GENEXPERT
        Case "KLITE":           Call TCPRcvData_KLITE
        Case "PPC300N":         Call TCPRcvData_PPC300N
        Case "VISIONB":         Call Phase_TCP_VISION
    
    End Select
                
End Sub

Private Sub SerialRcvData_UROMETER720()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strRData()      As String

    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    
On Error GoTo RST
    
'  Call SetSQLData("RCV", strRcvBuf, "A")
    
    RcvBuffer = Replace(RcvBuffer, vbLf, "")
    strRData = Split(RcvBuffer, vbCr)
    blnID = False
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            '
            If Mid(strRcvBuf, 1, 2) = "SD" Then
                'mResult 초기화
                Call SetResultBufFree
                'mOrder  초기화
                Call SetOrderBufFree
                
                blnID = True
                
                strSeq = Mid(strRcvBuf, 10)
                strSeq = Replace(strSeq, ")", "")
                strSeq = Replace(strSeq, "(", "")
                strSeq = Val(Trim(strSeq))
                
                '-- 결과정보
                mResult.Seq = strSeq
                'mResult.BarNo = strSeq
                With mResult
                    .Seq = strSeq
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- 결과환자정보
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
            End If
            
            If blnID = True Then
                strIntBase = Mid(strRcvBuf, 1, 4)
                strIntBase = Trim(strIntBase)
                strResult = ""
                strIntResult = ""
                strResult = Mid(strRcvBuf, 8, 4) '-- 정성
                strResult = Trim(strResult)
                strIntResult = strResult
                
                If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                    strIntResult = Trim(Mid(strRcvBuf, 4))  '-- 정량
                    strIntResult = Replace(strIntResult, "mg/dl", "")
                    strIntResult = Replace(strIntResult, "RBC/ul", "")
                    strIntResult = Replace(strIntResult, "WBC/ul", "")
                    
                    'strIntResult = Replace(strIntResult, "<", "")
                    'strIntResult = Replace(strIntResult, ">", "")
                    'strIntResult = Replace(strIntResult, "=", "")
                    strResult = strIntResult
                End If
                
                '-- 검사마스터 정보 가져오기
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                
                .spdResult.RowHeight(-1) = gROWHEIGHT

            End If
        Next
        
        '(%WBC% * %NEUT%) / 100
        Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)
            Call SetUpdateStatus(spdOrder, gRow, Res)
            strState = ""
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_UROMETER720" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_AFIAS2()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    Dim Res             As Integer
    
On Error GoTo RST
    
    If Mid(RcvBuffer, 1, 2) <> "$1" Then
        Exit Sub
    End If
    
    strRcvBuf = RcvBuffer
    With frmInterface
        '로그기록
        Call SetSQLData("RCV", strRcvBuf, "A")
        'mResult 초기화
        Call SetResultBufFree
        'mOrder  초기화
        Call SetOrderBufFree
        'Patient ID
        strBarno = Trim(mGetP(RcvBuffer, 5, "|"))
            
        '-- 결과정보
        With mResult
            .BarNo = strBarno
            .RsltDate = Format(Now, "yyyy-mm-dd")
            .RsltTime = Format(Now, "hh:mm:ss")
            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
        End With
        
        '-- 결과환자정보
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        If gRow <= 0 Then
            Exit Sub
        End If
        
        strState = "O"
        'Test Name (Channel)
        strIntBase = Trim(mGetP(strRcvBuf, 8, "|"))
        'Lot Name(for QC)
        'strLot = Trim(mGetP(strRcvBuf, 9, "|"))
        
        strResult = mGetP(strRcvBuf, 11, "|")
        
        strIntResult = strResult
                
        '-- 검사마스터 정보 가져오기
        If strIntBase <> "" And strResult <> "" Then
            If strState = "" Or strState = "O" Then
                strState = ""
            End If
            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                strState = "R"
            Else
                If strState = "" Then
                    strState = ""
                End If
            End If
        End If
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
    
        '(%WBC% * %NEUT%) / 100
        Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)
            Call SetUpdateStatus(spdOrder, gRow, Res)
            strState = ""
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AFIAS6" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_AFIAS6()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    Dim Res             As Integer
    
On Error GoTo RST
    
    '$1|FPRR020ND077|20170829005211|admin|03495843|||PCT|PCNCA16F|2018.12.26|< 0.10|||||ng/ml||||||||A2|||||||||
    
    '$1|A6RLG100006|20150120100700|CHOI|JOHN|30|M|PSA|PSLYC69|2017.06.26|12.45|||||ng/mL||||||G|G|0|||||
    '$1|A6RLG100006|20150120100700|CHOI|JOHN|30|M|HbA1C|HBLYC01|2017.12.31|5.1|20.22| mmol/mol |mg/dL||||G|G|

    If Mid(RcvBuffer, 1, 2) <> "$1" Then
        Exit Sub
    End If
    
    strRcvBuf = RcvBuffer
    With frmInterface
        '로그기록
        Call SetSQLData("RCV", strRcvBuf, "A")
        'mResult 초기화
        Call SetResultBufFree
        'mOrder  초기화
        Call SetOrderBufFree
        'Patient ID
        strBarno = Trim(mGetP(RcvBuffer, 5, "|"))
            
        '-- 결과정보
        With mResult
            .BarNo = strBarno
            .RsltDate = Format(Now, "yyyy-mm-dd")
            .RsltTime = Format(Now, "hh:mm:ss")
            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
        End With
        
        '-- 결과환자정보
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        If gRow <= 0 Then
            Exit Sub
        End If
        
        strState = "O"
        'Test Name (Channel)
        strIntBase = Trim(mGetP(strRcvBuf, 8, "|"))
        'Lot Name(for QC)
        'strLot = Trim(mGetP(strRcvBuf, 9, "|"))
        
        strResult = mGetP(strRcvBuf, 11, "|")
        
        '2021-02-17 장비단위 변경으로 10배 나누어 처리함
'''        If IsNumeric(strResult) Then
'''            strResult = strResult / 10
'''        Else
'''            If InStr(strResult, "<") > 0 Then
'''                strResult = Trim(Replace(strResult, "<", ""))
'''                strResult = strResult / 10
'''                strResult = "< " & strResult
'''            End If
'''
'''            If InStr(strResult, ">") > 0 Then
'''                strResult = Trim(Replace(strResult, ">", ""))
'''                strResult = strResult / 10
'''                strResult = "> " & strResult
'''            End If
'''        End If
        
        strIntResult = strResult
                
        'strIntResult = Replace(strIntResult, "<", "")
        'strIntResult = Replace(strIntResult, ">", "")
        'strIntResult = Replace(strIntResult, "=", "")
            
        '-- 검사마스터 정보 가져오기
        If strIntBase <> "" And strResult <> "" Then
            If strState = "" Or strState = "O" Then
                strState = ""
            End If
            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                strState = "R"
            Else
                If strState = "" Then
                    strState = ""
                End If
            End If
        End If
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
    
        '(%WBC% * %NEUT%) / 100
        Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
        
        .spdResult.RowHeight(-1) = gROWHEIGHT
        
        '## DB에 결과저장
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)
            Call SetUpdateStatus(spdOrder, gRow, Res)
            strState = ""
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AFIAS6" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_AFINION2()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    Dim Res             As Integer
    
    Dim intCnt          As Integer
    Dim strRData()      As String
    
On Error GoTo RST
    
    ReDim Preserve strRData(UBound(strRecvData))
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 0 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            '로그기록
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = Mid$(strRcvBuf, 1, 1)
            If IsNumeric(strType) Then
                strType = Mid$(strRcvBuf, 2, 1)
            End If
            Select Case strType
                Case "H"    '## Header
                    'mResult 초기화
                    Call SetResultBufFree
                    'mOrder  초기화
                    Call SetOrderBufFree
                Case "P"    '## Patient
                    strBarno = mGetP(strRcvBuf, 4, "|")
                Case "O"    '## Order
                    If strBarno = "" Then
                        strBarno = mGetP(strRcvBuf, 4, "|")
                    End If
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"

                Case "R"    '## Result
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                        
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = gROWHEIGHT
                Case "L"
                    '(%WBC% * %NEUT%) / 100
                    Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
                    
                    .spdResult.RowHeight(-1) = gROWHEIGHT
                    
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
                        Call SetUpdateStatus(spdOrder, gRow, Res)
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_AFINION2" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    
    Call SetSQLData("에러", "TCPRcvData_AFINION2" & vbCrLf & strErrMsg, "A")
    
    'frmErrMsg.txtErr = vbNewLine & strErrMsg
    'frmErrMsg.Show

End Sub

Public Sub FileData_APEX()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntName      As String
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    Dim i               As Integer
    Dim blnSame         As Boolean
    Dim blnID           As Boolean
    Dim Res             As Integer
    
    Dim intCnt          As Integer
    Dim strRData()      As String
    
    Dim r               As Integer
    Dim Buffer          As String
    Dim DeviceName      As String
    Dim DriverName      As String
    Dim PrinterPort     As String
    Dim PrinterName     As String
    
    Dim strDestFile     As String
    Dim strSrcfile      As String
    Dim k               As Integer
    
'On Error GoTo RST
    
    Screen.MousePointer = 11
    
   ' spdOrder.Visible = False
   ' spdResult.Visible = False
    
    If UBound(strRecvData) > 0 Then
         '-- 기본 프린터 변경
         Buffer = Space(1024)
         PrinterName = gB4C.PRTNAME '"Zan Image Printer(color)"
         r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
         'Parse the driver name and port name out of thebuffer
         GetDriverAndPort Buffer, DriverName, PrinterPort
         'MsgBox PrinterName
         'MsgBox DriverName
         'MsgBox PrinterPort
         If DriverName <> "" And PrinterPort <> "" Then
            SetDefaultPrinter PrinterName, DriverName, PrinterPort
         End If
    End If
    
    
    ReDim Preserve strRData(UBound(strRecvData))
    strRData = strRecvData
    
    '-- 프로그레스바 열기
    frmProgress.Show
    frmProgress.ZOrder 0
    frmProgress.Xprog.Min = 0
    frmProgress.Xprog.Max = UBound(strRecvData)
    
    With frmInterface
        
        For intCnt = 0 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            '로그기록
            Call SetSQLData("RCV", strRcvBuf, "A")
            strType = Mid$(strRcvBuf, 1, 1)
            If IsNumeric(strType) Then
                strType = Mid$(strRcvBuf, 2, 1)
            End If
            Select Case strType
                Case "H"    '## Header
                    'mResult 초기화
                    Call SetResultBufFree
                    'mOrder  초기화
                    Call SetOrderBufFree
                Case "P"    '## Patient
                    strBarno = mGetP(strRcvBuf, 2, "|")
                    strTubePos = mGetP(strRcvBuf, 3, "|")
                Case "O"    '## Order
                    strRackNo = mGetP(strRcvBuf, 2, "|")
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"

                Case "R"    '## Result
                    strIntName = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                    strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                    strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 3, "^"))
                    strIntResult = strResult
                    strFlag = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 4, "^"))
                        
                    If strIntBase = "tIgE" Then
                        If IsNumeric(strResult) Then
                            If strResult > 100 Then
                                strResult = ">100"
                            Else
                                strResult = "≤100"
                            End If
                        Else
                            If InStr(strResult, "2000") > 0 Then   '>2000
                                strResult = ">100"
                            ElseIf InStr(strResult, "<0.15") > 0 Then '<0.15
                                strResult = "≤100"
                            End If
                        End If
                    End If
                
                    '-- 검사마스터 정보 가져오기
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess_Allergy(mResult.BarNo, strIntName, strIntBase, strResult, strIntResult, strFlag) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = gROWHEIGHT - 5
                
                Case "L"
                    '(%WBC% * %NEUT%) / 100
                    'Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
                    
                    '.spdResult.RowHeight(-1) = gROWHEIGHT
                    
                    '-- 검체번호와 이름이 없으면 진행 안함
                    If mOrder.BarNo <> "" And mOrder.PNAME <> "" Then
                        '출력 스프레드 초기화
                        Call SetSpreadPrtClear
                        
                        '출력 스프레드에 데이터 뿌리기
                        Call SetSpreadPrtData
                        
                        '이미지 저장
                        .spdAllergy.PrintOrientation = PrintOrientationPortrait
                        .spdAllergy.Action = 13

                        '이미지 리네임(검체번호-이름)
                        '08-05-2021
                        strSrcfile = gB4C.IMAGE & Format(Now, "mm-dd-yyyy") & ".jpg"
                        
                        '대상 파일 이름을 정의
                        strDestFile = gB4C.IMAGE & Format(Now, "yyyy-mm-dd") & "/" & mOrder.BarNo & "-" & mOrder.PNAME & ".jpg"
                        
                        If Dir(gB4C.IMAGE & Format(Now, "yyyy-mm-dd"), vbDirectory) <> Format(Now, "yyyy-mm-dd") Then
                            Call MkDir(gB4C.IMAGE & Format(Now, "yyyy-mm-dd"))
                        End If
                            
                        Call Sleep(500)
                        
                        DoEvents
                        
                        'MsgBox strSrcfile
                        'MsgBox strDestFile
                        '원본을 대상에 복사
                        Call FileCopy(strSrcfile, strDestFile)
                        
                        Kill strSrcfile
                        
                        DoEvents
                        
                    End If
                    
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
                        Call SetUpdateStatus(spdOrder, gRow, Res)
                        strState = ""
                    End If
            End Select
            
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = intCnt
        
        Next
        
        spdOrder.Visible = True
        spdResult.Visible = True
        
        '-- 프로그레스바 닫기
        Unload frmProgress
        
        Screen.MousePointer = 0
    
        '원래 프린터로 되돌리기
        Buffer = Space(1024)
        PrinterName = gB4C.ORGPRTNAME '"Zan Image Printer(color)"

        r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))

        'Parse the driver name and port name out of thebuffer
        GetDriverAndPort Buffer, PrinterName, PrinterPort

        If DriverName <> "" And PrinterPort <> "" Then
           SetDefaultPrinter PrinterName, DriverName, PrinterPort
        End If
    
    End With

Exit Sub

RST:
    spdOrder.Visible = True
    spdResult.Visible = True
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_TCPRcvData_AFINION2" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SetSpreadPrtClear()
    Dim i           As Integer
    
    Call SetText(spdAllergy, "", 2, 2)
    Call SetText(spdAllergy, "", 3, 2)
    Call SetText(spdAllergy, "", 4, 2)
    Call SetText(spdAllergy, "", 5, 2)
    
    Call SetText(spdAllergy, "", 2, 8)
    Call SetText(spdAllergy, "", 3, 8)
    Call SetText(spdAllergy, "", 4, 8)
    Call SetText(spdAllergy, "", 5, 8)
    
    For i = 8 To 53
        Call SetText(spdAllergy, "", i, 4)
        Call SetText(spdAllergy, "", i, 5)
        Call SetText(spdAllergy, "", i, 10)
        Call SetText(spdAllergy, "", i, 11)
    Next
    
End Sub

Private Sub SetSpreadPrtData()
    Dim i           As Integer
    Dim blnData     As Boolean
    
    blnData = False
    
    With spdAllergy
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colPNAME), 2, 2)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colPSEX) & "/" & GetText(spdOrder, gRow, colPAGE), 3, 2)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colBARCODE), 4, 2)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colCHARTNO), 5, 2)
        
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colSPECIMEN), 2, 8)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colDEPT), 3, 8)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colHOSPDATE), 4, 8)
        Call SetText(spdAllergy, GetText(spdOrder, gRow, colEXAMDATE), 5, 8)
        
        SQL = ""
        SQL = SQL & "Select EQUIPCODE, RESULT, REFFLAG                                  " & vbCrLf
        SQL = SQL & "  From PATRESULT                                                   " & vbCrLf
        SQL = SQL & " Where EXAMDATE = '" & GetText(spdOrder, gRow, colEXAMDATE) & "'   " & vbCrLf
        SQL = SQL & "   And SAVESEQ  = " & GetText(spdOrder, gRow, colSAVESEQ) & vbCrLf
        SQL = SQL & "   And BARCODE  = '" & GetText(spdOrder, gRow, colBARCODE) & "'" & vbCrLf
        SQL = SQL & " Order By SEQNO "
        Set AdoRs_Local = New ADODB.Recordset
        AdoRs_Local.CursorLocation = adUseClient
        AdoRs_Local.Open SQL, AdoCn_Local
        If AdoRs_Local.RecordCount > 0 Then
            AdoRs_Local.MoveFirst
        End If
        If Not AdoRs_Local.EOF Then
            Do Until AdoRs_Local.EOF
                For i = 8 To 53
                    If Trim(UCase(GetText(spdAllergy, i, 3))) = Trim(UCase(AdoRs_Local("EQUIPCODE").Value)) Then
                        Call SetText(spdAllergy, AdoRs_Local("RESULT").Value, i, 4)
                        Call SetText(spdAllergy, AdoRs_Local("REFFLAG").Value, i, 5)
                        blnData = True
                        Exit For
                    End If
                Next
                If blnData = False Then
                    For i = 8 To 53
                        If Trim(UCase(GetText(spdAllergy, i, 9))) = Trim(UCase(AdoRs_Local("EQUIPCODE").Value)) Then
                            Call SetText(spdAllergy, AdoRs_Local("RESULT").Value, i, 10)
                            Call SetText(spdAllergy, AdoRs_Local("REFFLAG").Value, i, 11)
                            blnData = True
                            Exit For
                        End If
                    Next
                End If
                
                AdoRs_Local.MoveNext
                blnData = False
            Loop
            AdoRs_Local.Close
            Set AdoRs_Local = Nothing
        End If
        
        '검사자
        Call SetText(spdAllergy, gHOSP.USERNM, 64, 9)
        '의사
        Call SetText(spdAllergy, gHOSP.DOCNM, 65, 9)
        
            
        
    End With
    
End Sub


Private Sub SerialRcvData_HORIBA()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    
On Error GoTo RST

    strRData = Split(RcvBuffer, vbCr)
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            Select Case intCnt
                Case 4
                    If InStr(strRcvBuf, "AUTO_SID") > 0 Then
                        strSeq = Mid(strRcvBuf, InStr(strRcvBuf, "AUTO_SID") + 8)
                    Else
                        strSeq = mGetP(strRcvBuf, 2, Space(1))
                        strSeq = Val(strSeq)
                    End If
                    
                    '-- 결과정보
                    mOrder.Seq = strSeq
                    
                    With mResult
                        .BarNo = strSeq
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case 9 To 27
                    strIntBase = Trim(Mid(strRcvBuf, 1, 2))
                    strResult = Trim(Mid(strRcvBuf, 3))
                    strResult = Replace(strResult, "h", "")
                    strResult = Replace(strResult, "H", "")
                    strResult = Replace(strResult, "l", "")
                    strResult = Replace(strResult, "L", "")
                    strResult = Replace(strResult, " ", "")
                    
                    strResult = Replace(strResult, "S", "")
                    strResult = Replace(strResult, "s", "")
                    strIntResult = strResult
                    
                    If strIntBase = "'" Then
                        strIntBase = "|"
                    End If
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And strResult <> "" Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    spdOrder.RowHeight(-1) = gROWHEIGHT

                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_HORIBA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-- 장비설정
'System / System Condition /
'   [Test Requisition]
'       Routine:  BARCODE
'   [S.ID Barcode]
'       Barcode Type    : Multi
'       Digits          : 10
'       Check Mode      : No(No Chk.Chr.)
'System / Format /
'   Sample ID   Digits  : 20
Private Sub SerialRcvData_AU480()
    Dim RS_L            As ADODB.Recordset
    
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strAMRResult    As String   '수신한 결과(정성)
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim strCheck        As String   '검사오더체크
    Dim strSeqNo        As String   '검사순번
    Dim strOrderCode    As String   '처방코드
    Dim strTestName     As String   '검사코드
    Dim strAbbrName     As String   '검사코드
    Dim strTestCode     As String   '검사코드
    Dim strTestCodeSub  As String   '검사코드SUB
    Dim intResPrecUse   As Integer  '소수점변환여부
    Dim intResPrec      As Integer  '소수점자리수
    Dim strResType      As String   '소수점변환포맷
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '결과판정
    Dim strPrevRslt     As String   '이전결과
    
    Dim intRstRow       As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            
            Call SetSQLData("RCV", strRcvBuf, "A")
            
            strType = Mid$(strRcvBuf, 1, 2)

            Select Case strType
                Case "R "    '## Inquiry Order
                    'R 003201 0018          1013001917
                    'S 003201 0018          1013001917    E      13
                    
                    strBarno = Trim(Mid(strRcvBuf, 14, 20))
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = Mid(strRcvBuf, 9, 5)
                    End With
                    
                    Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        
                Case "D "    '## Result
                    'D 000103 0003          1908130030    E107  2.35  
                    
                    strBarno = Trim$(Mid$(strRcvBuf, 14, 20))
                    mResult.BarNo = strBarno
                    
                    '-- 결과정보
                    With mResult
                        .BarNo = strBarno
                        .RackNo = Mid(strRcvBuf, 3, 4)
                        .TubePos = Mid(strRcvBuf, 7, 2)
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    If strBarno = "" Then Exit Sub
    
                    strTmp = Mid$(strRcvBuf, 39)
                                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                
                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                
                        '-- 검사결과처리 프로세스
                        If strIntBase <> "" And strResult <> "" Then
                            If strState = "" Or strState = "O" Then
                                strState = ""
                            End If
                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                                strState = "R"
                            Else
                                If strState = "" Then
                                    strState = ""
                                End If
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop
                
                    spdOrder.RowHeight(-1) = gROWHEIGHT
                    

                    '## DB에 결과저장
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "저장실패", gRow, colSTATE
                        Else
                            '-- 저장 성공
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_AU480" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub Phase_Serial_HITACHI7180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        intPhase = 1
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_HITACHI7180
                        
                    Case vbCr
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_RP500()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim strSndData  As String
    
    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        
        Select Case BufChar
            Case STX
                AckOn = False
                RcvBuffer = BufChar
            Case EOT
                If AckOn = False Then
                    strSndData = STX & ACK & ETX & "0B" & EOT       'Ack Message
                    
                    Call SendData(strSndData)
                    
                    Call SerialRcvData_RP500
                End If
            Case ACK
                AckOn = True
                RcvBuffer = RcvBuffer & BufChar
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    
    Next i

End Sub

Private Sub Phase_Serial_HITACHI7020()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        intPhase = 1
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_HITACHI7020
                        
                    Case vbCr
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_SELEXON()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case ACK
            Case vbCr
            Case vbLf
            
                If strState = "ST" Then
                    tmrSelexOn.Interval = 1000
                    tmrSelexOn.Enabled = True
                End If
                If strState = "SI" Then
                    Call SerialRcvData_SELEXON
                    
                    tmrSelexOn.Interval = 1000
                    tmrSelexOn.Enabled = True
                End If
                If strState = "SE" Then
                    RcvBuffer = ""
                    strState = ""
                End If
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
                        
    Next i
    
End Sub

Private Sub Phase_Serial_DOTTO2000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_DOTTO2000
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_DOTTO2000
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub

Private Sub SerialRcvData_DOTTO2000()
    '장비 수신 변수
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정성)
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    '마스터 변수
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    '계산식 관련
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    Dim strCalChannel   As String
    Dim strCalCulate    As String
    Dim varCalCulate    As Variant
    Dim strCalNm(10)    As String
    Dim strCalCon(10)   As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            Call SetSQLData("RCV", strRcvBuf, "A")
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
        
            Select Case strType
                Case "H"    '## Header
                    strState = "H"
                Case "Q"    '## Request Information
                    '2Q|1|^1||ALL||||||||O
                    'strBarno = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                    strSeq = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))
                    
                    With mOrder
                        .Seq = strSeq
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_DOTTO2000(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    
                Case "P"    '## Patient
                    '2P|1||1||200528008|||N||||||0^Y
                    strBarno = mGetP(strRcvBuf, 6, "|")
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- 결과환자정보
                    Call SetPatInfo(strBarno, "0")
                    
                    strState = "P"
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                Case "O"
                    '3O|1|1^0001^^^|SAGES 200^||S|20200529101941|||||||||1||||||||||O
                    strState = "O"
                    
                Case "R"
                    '4R|1|^^^GPT|13|IU/L|0^40|N||F||||82009130154
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    strResult = strIntResult
                    
                    '-- 검사결과처리 프로세스
                    If strIntBase <> "" And (strIntResult <> "" Or strResult <> "") Then
                        If strState = "" Or strState = "O" Then
                            strState = ""
                        End If
                        If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                            strState = "R"
                        Else
                            If strState = "" Then
                                strState = ""
                            End If
                        End If
                    End If
                    
                    .spdResult.RowHeight(-1) = 15
                
                Case "C"
                Case "L"
                            '(%WBC% * %NEUT%) / 100
                    Call CalculateTest(mOrder.BarNo, gRow, spdOrder)
                    
                    .spdResult.RowHeight(-1) = gROWHEIGHT
                    
                    '## DB에 결과저장
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
                        Call SetUpdateStatus(spdOrder, gRow, Res)
                        strState = ""
                    End If

            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_DOTTO2000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub Phase_Serial_UROMETER720()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case "~"
                        RcvBuffer = ""
                        RcvBuffer = RcvBuffer & BufChar
                        intPhase = 2
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
            Case 2
            
                Select Case BufChar
                    Case "~"
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_UROMETER720
                        RcvBuffer = ""
                        intPhase = 1
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_UROMETER120()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "~"
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                
                Call SerialRcvData_UROMETER720
                'Call RcvData
                RcvBuffer = ""
                intPhase = 1
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_AFIAS2()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    '$1|FPRR020ND077|20170829005211|admin|03495843|||PCT|PCNCA16F|2018.12.26|< 0.10|||||ng/ml||||||||A2|||||||||
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                RcvBuffer = ""
                RcvBuffer = RcvBuffer & BufChar
            Case vbLf
            Case vbCr
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                Call SerialRcvData_AFIAS2
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_AFIAS6()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    '$1|FPRR020ND077|20170829005211|admin|03495843|||PCT|PCNCA16F|2018.12.26|< 0.10|||||ng/ml||||||||A2|||||||||
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                RcvBuffer = ""
                RcvBuffer = RcvBuffer & BufChar
            Case vbLf
            Case vbCr
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                Call SerialRcvData_AFIAS6
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_URINSCAN()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        RcvBuffer = ""
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case ETX
                        intPhase = 1
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_URINSCAN
                        RcvBuffer = ""
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i
End Sub

Private Sub Phase_Serial_AVL9180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
            
            Case ETX
                RcvBuffer = ""
            
            Case vbLf
                Call SerialRcvData_AVL9180
                
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_XP300()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            'Call SendOrder_XP300
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_XP300
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_YUMIZEN()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_YUMIZEN
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_YUMIZEN
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub


Private Sub Phase_TCP_XP300()
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                intFrameNo = intFrameNo + 1
                RcvBuffer = RcvBuffer & BufChar
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
    If InStr(RcvBuffer, "L|1|N") > 0 Then
        intPhase = 1
        intBufCnt = 0
        
        Call SerialRcvData_XP300
        
        intFrameNo = 0
        
    End If
    
End Sub

Private Sub Phase_TCP_YUMIZEN()
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                intFrameNo = intFrameNo + 1
                RcvBuffer = RcvBuffer & BufChar
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
    If InStr(RcvBuffer, "L|1|N") > 0 Then
        intPhase = 1
        intBufCnt = 0
        
        Call SerialRcvData_YUMIZEN
        
        intFrameNo = 0
        
    End If
    
End Sub

Private Sub Phase_Serial_EPOC()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
            Case "-"
                RcvBuffer = RcvBuffer & BufChar
                If InStr(RcvBuffer, "----------------------") > 0 Then
                    Call SerialRcvData_EPOC
                    RcvBuffer = ""
                End If
            Case ETX
            Case EOT
                Call SerialRcvData_EPOC
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
        
End Sub



Private Sub Phase_Serial_ISMART30()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            '
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_ISMART30
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_MICROS60()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            '
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_MICROS60
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_XL1000I()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                    RcvBuffer = ""
            Case ETB
                    Call SendData(ACK)
            Case ETX
                    Call SerialRcvData_XL1000I
                    RcvBuffer = ""
            Case Else
                    RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_LTC52()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                    RcvBuffer = BufChar
            Case ETX
                    Call SerialRcvData_LTC52
                    RcvBuffer = ""
            Case Else
                    RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_ARKRAY()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        'If strState = "Q" Then
                            'Call SendOrder_XP300
                        'End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX

                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                        
                    Case ETX
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                        If blnIsETB = False Then
                            intPhase = 4
                        Else
                            intPhase = 2
                        End If
                                
                        Call SendData(ACK)
                    
                    Case vbLf
                        'intPhase = 4
                        'Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_ARKRAY
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_BIOLYTE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

'On Error Resume Next

 '   Debug.Print pBuffer
    
    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                        RcvBuffer = ""
                        tmrBIOLYTE.Interval = 3000
                        tmrBIOLYTE.Enabled = True
'                    Case ACK
'                        'If strState = "Q" Then
'                            'Call SendOrder_XP300
'                        'End If
                End Select
            Case 2      '## Transfer Phase
                
                    
                Select Case BufChar
'                    Case ENQ
'                        Erase strRecvData
'                        Call SendData(ACK)
'                    Case STX
'                        If intBufCnt = 0 Then
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                        Else
'                            If Len(strRecvData(intBufCnt)) > 5 Then
'                                intBufCnt = intBufCnt + 1
'                                ReDim Preserve strRecvData(intBufCnt)
'                            End If
'                        End If
''                    Case ETB
''                        blnIsETB = True
''                        intPhase = 3
''                    Case ETX
''                        'intBufCnt = intBufCnt + 1
''                        'ReDim Preserve strRecvData(intBufCnt)
''                        'intPhase = 3
'                    Case vbCr
                    Case vbLf
                        Call SendData(ACK)
'
''                    Case EOT
''                        intPhase = 1
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar

                End Select

        End Select
    Next i
    
End Sub


Private Sub SendOrder_STAGO()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            
            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                intSndPhase = 3
            Else
                intSndPhase = 2
            End If

            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|||" & mOrder.PID & "|^1^1^56|||19700505" & vbCr & ETX
            intSndPhase = 4
            intFrameNo = intFrameNo + 1

        Case 3  '## No Order
            strOutput = intFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            intSndPhase = 5

        Case 4  '## Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 4
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 5
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 4
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 5
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1

        Case 6  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_THUNDERBOLT()
    Dim strOutput   As String     '송신할 데이터
    Dim blnLast     As Boolean
    Dim intRow      As Integer
    Dim strBarno    As String
    Dim strItems    As String
    Dim varItem     As Variant
    Dim i           As Integer
    Dim strTmp      As String
    
    blnLast = False

    With spdOrder
        If intSndPhase <= 3 Then
            For intRow = 1 To .DataRowCnt
                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "오더준비" Then
                    strBarno = Trim(GetText(spdOrder, intRow, colBARCODE))
                    strItems = Trim(GetText(spdOrder, intRow, colSPECIMEN))
                    If intSndPhase = 3 Then
                        varItem = Split(strItems, "@")
                        If UBound(varItem) > 0 Then
                            strItems = varItem(0)
                            
                            For i = 1 To UBound(varItem)
                                strTmp = strTmp & "@" & varItem(i)
                            Next
                            strTmp = Mid(strTmp, 2)
                            Call SetText(spdOrder, strTmp, intRow, colSPECIMEN)
                        Else
                            Call SetText(spdOrder, "0", intRow, colCHECKBOX)
                            Call SetText(spdOrder, "오더전송", intRow, colSTATE)
                            
                            If intRow = .DataRowCnt Then
                                blnLast = True
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next
        End If
    End With
    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||LIS|||||||P|LIS2-A2|" & Format(Now, "yyyyMMddHHmmss") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|" & mPNo & "||" & strBarno & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
            mPNo = mPNo + 1
        
        Case 3  '## Order
            strOutput = intFrameNo & "O|" & mOCnt & "|" & strBarno & "||" & strItems & "|R" & vbCr & ETX
            If blnLast = True Then
                intSndPhase = 4
            Else
                If UBound(varItem) > 0 Then
                    mOCnt = mOCnt + 1
                    intSndPhase = 3
                Else
                    mOCnt = 1
                    intSndPhase = 2
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            strQState = ""
            Call SendData(EOT)
            intFrameNo = 1
            mPNo = 1
            mOCnt = 1
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_XN1000()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            strQState = ""
            Call SendData(EOT)
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800_ASTM()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            '<STX>1                   H|\^&|||HostName^^^^|||||CA-600<CR><ETX><CHK1><CHK2><CR><LF
            strOutput = intFrameNo & "H|\^&|||HostName^^^^|||||CA-600" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmInterface.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800()
    Dim strOutput   As String     '송신할 데이터

    Select Case intSndPhase
        Case 1  '## Header
            '<STX>1                   H|\^&|||HostName^^^^|||||CA-600<CR><ETX><CHK1><CHK2><CR><LF
            strOutput = intFrameNo & "H|\^&|||HostName^^^^|||||CA-600" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N" & vbCr & ETX
                intSndPhase = 4
            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmInterface.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


Private Sub SendOrder_ACCESS2()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer
    Dim intDestRow  As Integer

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|" & mOrder.PID & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## No Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                'strOutput = "O|1|" & mOrder.BarNo & "|" & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "|" & mOrder.Order & "|R||||||A||||" & "Serum"
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & "Serum"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            intSndPhase = 1
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)
    
End Sub

Private Sub SendOrder_ACCESS2_Batch()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer
    Dim intDestRow  As Integer
    Dim blnOrder    As Boolean
    Dim blnLast     As Boolean

    blnOrder = False
    blnLast = True
    
    With spdOrder
        If intSndPhase = 2 Or intSndPhase = 3 Then
            For intRow = 1 To .MaxRows
                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
                    mOrder.BarNo = Trim(GetText(spdOrder, intRow, colBARCODE))
                    mOrder.PID = Trim(GetText(spdOrder, intRow, colPID))
                    mOrder.RackNo = Trim(GetText(spdOrder, intRow, colRACKNO))
                    mOrder.TubePos = Trim(GetText(spdOrder, intRow, colPOSNO))
                    'mOrder.Order = Trim(GetText(spdOrder, intRow, colDEPT))
                    mOrder.Order = Trim(GetTag(spdOrder, intRow, colSTATE))
                    mOrder.DestRow = intRow
                    'blnOrder = True
                    'intDestRow = intRow
                    Exit For
                End If
            Next
'            For intRow = intDestRow + 1 To .MaxRows
'                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
'                    blnLast = False
'                    Exit For
'                End If
'            Next
        End If
    End With
    
    If blnOrder = True Then
        Select Case intSndPhase
            Case 1  '## Header
                strOutput = intFrameNo & "H|\^&|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
                intSndPhase = 2
                intFrameNo = intFrameNo + 1
                
            Case 2  '## Patient
                strOutput = intFrameNo & "P|1|" & mOrder.PID & vbCr & ETX
                intSndPhase = 3
                intFrameNo = intFrameNo + 1
    
            Case 3  '## No Order
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    'strOutput = "O|1|" & mOrder.BarNo & "|" & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "|" & mOrder.Order & "|R||||||A||||" & "Serum"
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||" & "Serum"
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## 남은 문자열이 있을때
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
                intFrameNo = intFrameNo + 1
    
            Case 4  '## Termianator
                strOutput = intFrameNo & "L|1|N" & vbCr & ETX
                intSndPhase = 5
                intFrameNo = intFrameNo + 1
    
            Case 5  '## EOT
                strState = ""
                Call SendData(EOT)
                intFrameNo = 1
                intSndPhase = 1
                
                Call SetText(spdOrder, "0", mOrder.DestRow, colCHECKBOX)
                Call SetText(spdOrder, "오더전송", mOrder.DestRow, colSTATE)
                
                blnLast = True
                For intRow = mOrder.DestRow + 1 To spdOrder.MaxRows
                    If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "" Then
                        blnLast = False
                        Exit For
                    End If
                Next

                If blnLast = False Then
                    strState = "Q"
                    Call SendData(ENQ)
                End If
                Exit Sub
        End Select
    
        If intFrameNo = 8 Then
            intFrameNo = 0
        End If
    
        strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
        Call SendData(strOutput)
    End If
    
End Sub


Private Sub SendOrder_DOTTO2000()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer
    Dim strBarcode  As String
    Dim strSeq      As String
    Dim strPID      As String
    Dim strPName    As String
    Dim strPSex     As String
    Dim strPAge     As String
    Dim strAgeUnit  As String
    Dim strDisk     As String
    Dim strPos      As String
    Dim strOrder    As String
    
'    With spdOrder
'        If intSndPhase = 2 Or intSndPhase = 3 Then
'            For intRow = 1 To .DataRowCnt
'                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colBARCODE) <> "" And GetText(spdOrder, intRow, colSTATE) = "오더준비" Then
'                    strBarcode = GetText(spdOrder, intRow, colBARCODE)
'                    mOrder.BarNo = strBarcode
'                    strSeq = GetText(spdOrder, intRow, colSEQNO)
'                    strPName = "" 'GetText(spdOrder, intRow, colPNAME)
'                    strPSex = "" 'GetText(spdOrder, intRow, colPSEX)
'                    strPAge = "" 'GetText(spdOrder, intRow, colPAGE)
'                    strAgeUnit = "" 'Y , M , D
'                    strDisk = "" 'GetText(spdOrder, intRow, colRACKNO)
'                    strPos = "" 'GetText(spdOrder, intRow, colPOSNO)
'                    mOrder.Order = GetText(spdOrder, intRow, colSPECIMEN)
'                    Exit For
'                End If
'            Next
'        End If
'    End With
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            mOrder.SendCnt = 0
            
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||" & mOrder.BarNo & "||" & mOrder.PNAME & "|||" & strPSex & "||||||" & strPAge & "^" & strAgeUnit & vbCr & ETX
            strOutput = intFrameNo & "P|1||" & mOrder.BarNo & "||" & mOrder.PNAME & "|||||||||^" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## Order
            mOrder.SendCnt = mOrder.SendCnt + 1
            strOutput = mGetP(mOrder.Order, mOrder.SendCnt, "\")
            
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                                                                                                                                                                                           '1   - Serum
                                                                                                                                                                                           '2   - Urine
                                                                                                                                                                                           '3   - CSF
                                                                                                                                                                                           '4   - Suprnt
                                                                                                                                                                                           '5   - Others
                'strOutput = "O|" & mOrder.SendCnt & "|" & mOrder.BarNo & "^^" & mOrder.Seq & "^" & strDisk & "^" & strPos & "^N||" & strOutput & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||||||1||||||||||O"
                strOutput = "O|" & mOrder.SendCnt & "|" & mOrder.BarNo & "^^" & mOrder.Seq & "^^^N||" & strOutput & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||||||1||||||||||O"
                
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 3
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 3
                End If
            End If
            
            If mGetP(mOrder.Order, mOrder.SendCnt + 1, "\") = "" Then
                intSndPhase = 4
            End If
            
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            intSndPhase = 1
            With spdOrder
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = colBARCODE
                    If Trim(.Text) = mOrder.BarNo Then
                        Call SetText(spdOrder, "0", intRow, colCHECKBOX)
                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
                        Exit For
                    End If
                Next
            End With

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)

End Sub

Private Sub SendOrder_YUMIZEN()
    Dim strOutput   As String     '송신할 데이터
    Dim intRow      As Integer

    Select Case intSndPhase
        Case 1  '## Header
                                     'H|\^&|||HCM|||||||P|LIS2-A2|20150323160111<CR><ETX>51<CR><LF
            strOutput = intFrameNo & "H|\^&|||HCM|||||||P|LIS2-A2|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
            
        Case 2  '## Patient
                                     'P|1||2                 ||BOND^JAMES||19770526|M|||||<CR><ETX>24<CR><LF
            strOutput = intFrameNo & "P|1||" & mOrder.PID & "||^||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## Order
                                    '3O|1|289645146||^^^DIF|R|20150323160111|||||N||||||||||||||Q|||||<CR><ETX>C0<CR><LF
            '## 최초 보낼때
            If mOrder.IsSending = False Then
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "|||||N||||||||||||||Q|||||"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            '## 남은 문자열이 있을때
            Else
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 4
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
                                    '4L|1|<CR><ETX>B9<CR><LF
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            
            With spdOrder
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = colBARCODE
                    If Trim(.Text) = mOrder.BarNo Then
                        Call SetText(spdOrder, "오더전송", intRow, colSTATE)
                        Exit For
                    End If
                Next
            End With
            
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    Call SendData(strOutput)

End Sub

Private Sub Phase_Serial_STAGO()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_STAGO
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub


Private Sub Phase_Serial_MEDONIC()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
                        
    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1
                If BufChar = "<" Then
                    RcvBuffer = ""
                    RcvBuffer = RcvBuffer & BufChar
                    intPhase = 2
                End If
                
            Case 2
                
                RcvBuffer = RcvBuffer & BufChar
                
                If InStr(RcvBuffer, "End:Chksum") > 0 Then
'                    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                    
                    intPhase = 1
                    
'                    Call SerialRcvData_MEDONIC
                    
                    RcvBuffer = ""
               End If
        End Select
    Next i
                       
                       
End Sub

Private Sub Phase_Serial_XN1000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim ix1         As Integer
    
    lngBufLen = Len(pBuffer)
    
    For ix1 = 1 To lngBufLen
        BufChar = Mid$(pBuffer, ix1, 1)

        Select Case intPhase
            Case 1
                Select Case Asc(BufChar)
                    Case 5      'ENQ
                        intPhase = 2
                        
                        RstEnd = "Y"
                        bSTXChk = False
                        bEndChk = True
                        
                        Call SendData(ACK)

                    Case Else
                        intPhase = 1
                End Select

            Case 2
                Select Case Asc(BufChar)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     'LF
                        If bEndChk = True Then
                            Call SerialRcvData_XN1000
                            RcvBuffer = ""
                        End If
                        Call SendData(ACK)

                    Case 13     'CR
                        If bEndChk = True Then
                            Call SerialRcvData_XN1000
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If strState = "Q" Then
                            Call SendData(ENQ)
                            intSndPhase = 1
                        End If
                        intPhase = 3

                    Case 5      'ENQ
                        bSTXChk = True
                        bEndChk = True
                        Call SendData(ACK)

                    Case 21     'NAK
                        Call SerialRcvData_XN1000
                        
                        intSndPhase = 1
                        intFrameNo = 1

                        Call SendData(ENQ)

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & BufChar
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(BufChar)
                    Case 6      'ACK
                        If strState = "Q" Then
                            Call SendOrder_XN1000
                        End If

                    Case 5      'ENQ
                        bSTXChk = False
                        bEndChk = True
                        Call SendData(ACK)
                        intPhase = 2

                    Case 21     'NAK
                        intSndPhase = 1
                        intFrameNo = 1
                        Call SendData(ENQ)
                        intPhase = 3

                    Case 4      'EOT
                        intPhase = 1

                End Select
        End Select
    Next ix1
    
    
'''    For i = 1 To lngBufLen
'''        BufChar = Mid$(pBuffer, i, 1)
'''        Select Case intPhase
'''            Case 1      '## Estabilshment Phase
'''                Select Case BufChar
'''                    Case ENQ
'''                        Erase strRecvData
'''                        intPhase = 2
'''                        Call SendData(ACK)
'''                    Case ACK
'''                        If strState = "Q" Then
'''                            Call SendOrder_XN1000
'''                        End If
'''                End Select
'''            Case 2      '## Transfer Phase
'''                Select Case BufChar
'''                    Case ENQ
'''                        Erase strRecvData
'''                        Call SendData(ACK)
'''                    Case STX
'''                        If intBufCnt = 0 Then
'''                            intBufCnt = 1
'''                            Erase strRecvData
'''                            ReDim Preserve strRecvData(intBufCnt)
'''                        Else
'''                            intBufCnt = intBufCnt + 1
'''                            ReDim Preserve strRecvData(intBufCnt)
'''                        End If
'''                    Case ETB
'''                        blnIsETB = True
'''                        intPhase = 3
'''                    Case ETX
'''                        intBufCnt = intBufCnt + 1
'''                        ReDim Preserve strRecvData(intBufCnt)
'''                        intPhase = 3
'''                    Case vbCr
'''                    Case vbLf
'''                    Case EOT
'''                        intPhase = 1
'''                    Case Else
'''                        If blnIsETB = False Then
'''                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'''                        Else
'''                            blnIsETB = False
'''                        End If
'''                End Select
'''            Case 3      '## Transfer Phase
'''                Select Case BufChar
'''                    Case vbCr
'''                    Case vbLf
'''                        intPhase = 4
'''                        Call SendData(ACK)
'''                End Select
'''            Case 4      '## Termination Phase
'''                Select Case BufChar
'''                    Case STX
'''                        intPhase = 2
'''                    Case EOT
'''                        intPhase = 1
'''                        intBufCnt = 0
'''
'''                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
'''                        Call SerialRcvData_XN1000
'''
'''                        Erase strRecvData
'''
'''                        If strState = "Q" Then
'''                            intSndPhase = 1
'''                            intFrameNo = 1
'''                            Call SendData(ENQ)
'''                        End If
'''                End Select
'''        End Select
'''    Next i
    
    
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case cIF.Phase
'            Case 1      '## Estabilshment Phase
'                Select Case BufChar
'                    Case ENQ
'                        cIF.BufCnt = 1
'                        cIF.ClearBuffer
'                        Call SendData(ACK)
'                        cIF.Phase = 2
'                    Case ACK
'                        If cIF.State = "Q" Then
'                            Call SendOrder_XN1000
'                        Else
'                            Call SendData(ACK)
'                        End If
'                End Select
'
'            Case 2      '## Transfer Phase
'                Select Case BufChar
'                    Case ENQ
'                        cIF.BufCnt = 1
'                        cIF.ClearBuffer
'                        Call SendData(ACK)
'                    Case STX
'                    Case vbCr
'                        cIF.BufCnt = cIF.BufCnt + 1
'                    Case ETB
'                        cIF.IsETB = True
'                        cIF.Phase = 3
'                    Case ETX
'                        cIF.Phase = 3
'                    Case Else
'                        If cIF.IsETB = False Then
'                            Call cIF.AddBuffer(BufChar)
'                        Else
'                            cIF.IsETB = False
'                        End If
'                End Select
'
'            Case 3      '## Transfer Phase
'                Select Case BufChar
'                    Case vbCr
'                    Case vbLf
'                        If cIF.IsETB = False Then
'                            cIF.Phase = 4
'                        Else
'                            cIF.Phase = 2
'                        End If
'                        Call SendData(ACK)
'
'                End Select
'
'            Case 4      '## Termination Phase
'                Select Case BufChar
'                    Case STX
'                        cIF.Phase = 2
'                    Case EOT
'                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
'                        Call SerialRcvData_XN1000
'
'                        If cIF.State = "Q" Then
'                            cIF.SndPhase = 0
'                            cIF.FrameNo = 0
'                            Call SendData(ACK)
'                        End If
'
'                        cIF.Phase = 1
'
'                End Select
'        End Select
'    Next i
'
'
End Sub

Private Sub Phase_Serial_THUNDERBOLT()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        intPhase = 2

                        Erase strRecvData
                        Call SendData(ACK)

                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_THUNDERBOLT
                        End If

                End Select

            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If

                    Case ETB
                        blnIsETB = True
                        intPhase = 3

                    Case vbCr
                    Case vbLf
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case EOT
                        intPhase = 1

                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select

            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = IIf(blnIsETB = False, 4, 2)
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1

                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")

                        Call SerialRcvData_THUNDERBOLT
                        Erase strRecvData

                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If

                End Select
        End Select
    Next i


End Sub

Private Sub Phase_Serial_MINIVIDAS()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case EOT    '4
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")

                Call SerialRcvData_MINIVIDAS
                RcvBuffer = ""
            
            Case ENQ    '5
                RcvBuffer = ""
                Call SendData(ACK)  '6

            Case GS     '29
                RcvBuffer = ""
                Call SendData(ACK)

            Case Else
                RcvBuffer = RcvBuffer & BufChar

        End Select
    Next i

End Sub

Private Sub Phase_Serial_CA800_ASTM()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_CA800
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr, vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_CA800
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

    
End Sub

Private Sub Phase_Serial_CA800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
                
            Case ETX
                Call Sleep(200)         '0.2 sec or More Delay
                
                Call SendData(ACK)
                
                Call SerialRcvData_CA800
            
            Case ACK
            
            Case NAK
            
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i


    
End Sub

Private Sub Phase_Serial_ACCESS2()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_ACCESS2
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_ACCESS2
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_PPC300N()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                    
                        Call SerialRcvData_STAGO
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)

                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_PATHFAST()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'
'        Select Case BufChar
'            Case ENQ
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'                comEqp.Output = ACK
'                SetRawData "[Tx]" & ACK
'            Case STX
'                intBufCnt = intBufCnt + 1
'                ReDim Preserve strRecvData(intBufCnt)
'
'            Case vbLf
'                comEqp.Output = ACK
'                SetRawData "[Tx]" & ACK
'
'            Case EOT
'                dtpToday.Value = Now
'                Call SerialRcvData_PATHFAST
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendData(ACK)
                    Case ACK
                        If strState = "Q" Then
'                            Call SendOrder_PATHFAST
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        intBufCnt = 0
                        
                        Call SerialRcvData_PATHFAST
                        
                        Erase strRecvData
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                End Select
        End Select
    Next i
    
End Sub

Private Sub Phase_Serial_AU480()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            Case ETB
            Case ETX
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                
                Call SerialRcvData_AU480
                RcvBuffer = ""
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    
    Next i
    
End Sub


Private Sub Phase_Serial_HORIBA()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
                RcvBuffer = RcvBuffer & BufChar
            Case ETX
                MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                
                Call SerialRcvData_HORIBA
                RcvBuffer = ""
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
    
End Sub


Private Sub Phase_TCP_KLITE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_KLITE
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub

Private Sub Phase_TCP_PPC300N()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_PPC300N
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub

Private Sub Phase_TCP_GENEXPERT()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                        Call SendWSckData(ACK)
                End Select
            Case 2
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendWSckData(ACK)
                    Case STX
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case EOT
                        intPhase = 1
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendWSckData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        intPhase = 1
                        Call TCPRcvData_GENEXPERT
                        
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_TCP_AFINION2()
    
    pBuffer = Replace(pBuffer, vbLf, "")
    strRecvData = Split(pBuffer, vbCr)
                
    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    
    If picComm.Visible = False Then
        Call SendWSckData(ACK)
    End If
    Call TCPRcvData_AFINION2

End Sub


Private Sub Phase_TCP_VISION()
    Dim Buffer      As Variant
    Dim BufChar     As String
    'Dim lngBufLen   As Long
    Dim i           As Long

    Dim strBuffer   As String
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strRcvCnt   As String
    Dim strSendAck  As String

    Dim strNS       As String
    Dim strNE       As String
    Dim intNS       As Integer
    Dim intNE       As Integer

    Dim strSendData As String

    strRecvData = Split(pBuffer, vbLf)

    For i = 0 To UBound(strRecvData)
        strBuffer = strRecvData(i)
        If strBuffer = "" Then
            Exit For
        End If
        strLastSeq = mGetP(strBuffer, 1, vbTab)
        strRcvSign = mGetP(strBuffer, 2, vbTab)
        strSendAck = strLastSeq & vbTab & "ACK"

        Select Case UCase(strRcvSign)
            Case "RESULT"
                '2   RESULT  1   VC0111  2015-11-03T06:55:19Z    3   3   23.3    21  17  23.5625 24.8125 False   False
                '3   RESULT  2   VC0111  2015-11-03T06:55:19Z    4   4   24.0    96  84  23.5625 24.8125 False   False

                Call TCPRcvData_VISION
                strBuffer = ""

            Case "CONNECT"
                strSendData = strSendAck & vbLf

                Call SendWSckData(strSendData)

            Case "RESULTS"
                '결과요청
                strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))

                strNS = strRcvCnt
                strNE = mGetP(strBuffer, 4, vbTab)

                strNS = strNS - strNE
                strNE = strNS + strNE

                strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE & vbLf

                Call SendWSckData(strSendData)

                'Call WritePrivateProfileString("config", "LASTSEQ", strRcvCnt, App.PATH & "\Interface.ini")
                
                txtLastSeq.Text = strRcvCnt

                'blnResults = False
        End Select
    Next i


End Sub

Public Sub frmClear()
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    spdWork.MaxRows = 0
    
    'dtpFrom.Value = Now
    'dtpTo.Value = Now
    
    txtBarNum.Text = ""
    txtRackNo.Text = "1"
    txtPosNo.Text = "1"
    txtSeqNo.Text = "1"
    txtOldBarNum.Text = ""
    txtFrNo.Text = "0000"
    txtToNo.Text = "0999"
    
    lblBarcode.Caption = ""
    lblPatNm.Caption = ""
    lblStatus.Caption = ""
    
    'BIT Json
    lblSlipCd.Caption = gHOSP.PARTCD
    
    'UROMETER
    txtWBC.Text = ""
    txtRBC.Text = ""
    
    'VISION
    txtRCnt.Text = "1"
    


    'spdOrder.EnhanceStaticCells = True
        
    'clrSelectedHoverUpperColor Selected, active background column header color of the upper half of the header when the mouse pointer is over the header (default value is RGB(255,212,141)
    'clrSelectedHoverLowerColor Selected, active background column header color of the lower half of the header when the mouse pointer is over the header (default value is RGB(242,147,59)
    'clrSelectedUpperColor Selected, upper background color of the active column header (default value is RGB(249,217,159)
    'clrSelectedLowerColor Selected, lower background color of the active column header (default value is RGB(241,193,96)
    'clrHoverUpperColor Upper background color of the column header when the mouse pointer is over the header (default value is RGB(223,226,228)
    'clrHoverLowerColor Lower background color of the column header when the mouse pointer is over the header (default value is RGB(189,197,210)
    'clrUpperColor Background color of the upper half of the header (default value is RGB(249,252,253)
    'clrLowerColor Background color of the lower half of the header (default value is RGB(211,220,233)
    'clrSelectedBorderColor Border color of the column header when the column is selected (default value is RGB(242,149,54)
    'clrBorderColor Border color of the column header (default value is RGB(158,182,206)


    'Call spdOrder.SetEnhancedColumnHeaderColors(RGB(255, 212, 141), _
                                                RGB(242, 147, 59), _
                                                vbWhite, _
                                                vbGreen, _
                                                RGB(223, 226, 228), _
                                                RGB(189, 197, 210), _
                                                RGB(249, 252, 253), _
                                                RGB(211, 220, 233), _
                                                RGB(242, 149, 54), _
                                                RGB(158, 182, 206))


End Sub


Private Sub Command1_Click()
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    'If Row = 0 Then
        'If Col = colCHECKBOX Then
            If GetText(spdWork, 1, colCHECKBOX) = "1" Then
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "0", i, colCHECKBOX)
                Next
            Else
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "1", i, colCHECKBOX)
                Next
            End If
        'Else
            '-- 정렬 추가
        '    Call SetSpreadSort(spdWork, 0)
        'End If
        Exit Sub
    'End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        If MsgBox("인터페이스 화면을 닫으시겠습니까?", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    Dim strIFStatus As String
    Dim intCol      As Integer
    Dim i           As Integer
    
On Error GoTo ErrHandle
    
    
    spdOrder.GrayAreaBackColor = 16316655
    spdWork.GrayAreaBackColor = 16316655  'spdOrder.GrayAreaBackColor
    spdResult.GrayAreaBackColor = 16316655
    
    Me.Caption = gHOSP.MACHNM

    '-- 설정상태 표시
    If gHOSP.BARUSE = "Y" Then
        strIFStatus = "▣ 바코드사용"
    Else
        If gHOSP.RSTTYPE = "1" Then
            strIFStatus = "▣ 순번 맞춤"
        ElseIf gHOSP.RSTTYPE = "2" Then
            strIFStatus = "▣ R/P 맞춤"
        ElseIf gHOSP.RSTTYPE = "3" Then
            strIFStatus = "▣ 체크순"
        End If
    End If
    
    '-- 워크조회 위치
    If gWORKPOS = "M" Then
        spdWork.Visible = True
        fraWorkInfo.Visible = True
    Else
        spdWork.Visible = False
        fraWorkInfo.Visible = False
    End If

    '-- 폼 초기화
    Call frmClear
    
    Call SetSpreadPrtClear
    
    '-- 통신변수 초기화
    Call CtlInitializing
    
    '-- 메뉴 설정
    Call SetMenu

    '-- 컬럼헤더설정
    'Call SetColumnHeader(spdOrder)

    '-- 컬럼보이기설정
    Call SetColumnView(spdOrder)
    
    '-- 컬럼보이기설정
    Call SetColumnView(spdWork)
    
    '-- 컬럼보이기설정
    Call SetColumnViewResult(spdResult)
    
    '-- 워크리스트 항목 조정(FIX)
    For intCol = 1 To colSTATE
        spdWork.Col = intCol
        If intCol = colHOSPDATE Or intCol = colBARCODE Or intCol = colPNAME Or intCol = colPID Or intCol = colCHECKBOX Then
            spdWork.ColHidden = False
        Else
            spdWork.ColHidden = True
        End If
    Next
    
    '-- 검사마스터정보 gArrEQP(검사갯수,1~17) 에 담기
    Call GetTestList

    '-- 검사코드 gAllTestCd 에 담기
    Call GetTestCodeList

    '-- 메인 스프레드[spdOrder]에 검사명 보이기
    Call SetExamCode(spdOrder)

    '-- 통신 상태 관련 컨트롤 초기화
    MDIIF.imgNet1.ZOrder 0
    tmrDBConn.Interval = 1000
    tmrDBConn.Enabled = True
    tmrQ = False
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    '-- 이전결과 삭제
    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gHOSP.SAVEDAY + "일전 데이타를 삭제하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    
    'Me.Icon = LoadPicture(App.PATH & "\ICON\" & gHOSP.PARTNM & ".ico")
    'Me.Icon = LoadPicture(App.PATH & "\ICON\LFT.ico")
    
    'LFT 계산식 관련 채널 가져오기
    If gHOSP.PARTNM = "LFT" Then
        gTC = GetTestCh("TC")
        gTG = GetTestCh("TG")
        gHDL = GetTestCh("HDL")
        gLDLC = GetTestCh("LDLC")
        gBUN = GetTestCh("BUN")
        gCREA = GetTestCh("CREA")
        geGFR = GetTestCh("eGFR")
        gBCRatio = GetTestCh("BCRatio")
    End If
    
    If gHOSP.DBCONCHK = "Y" Then
        tmrConn.Interval = 60000
        'tmrConn.Interval = 5000
        tmrConn.Enabled = True
    Else
        tmrConn.Enabled = False
    End If
    
    '-- 의사랑일 경우 'XML 정리' 버튼 보이게함.
    If UCase(gEMR) = "UBCARE" Then
        MDIIF.cmdXML.Visible = True
    Else
        MDIIF.cmdXML.Visible = False
    End If
    
    
    '-- 변수 초기화(E-170/H-7600)
    RstEnd = "Y"
    bSTXChk = False
    bEndChk = True
    
    If gHOSP.MACHNM <> "EPOC" Then
        '-- 통신열기
        Call OpenCommunication
    End If
    
    If gHOSP.MACHNM = "EPOC" Then
        tmrEPOC.Interval = 10000
        tmrEPOC.Enabled = True
        'Call WinExec("C:\IF_EPOC\COM\EPOC.exe 5 C:\IF_EPOC\LOG", 0)
        'FileEPOC.PATH = "C:\IF_EPOC\Log"
        Select Case gComm.COMPORT
            Case "1":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 1 C:\IF_EPOC\COM\LOG", 0)
            Case "2":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 2 C:\IF_EPOC\COM\LOG", 0)
            Case "3":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 3 C:\IF_EPOC\COM\LOG", 0)
            Case "4":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 4 C:\IF_EPOC\COM\LOG", 0)
            Case "5":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 5 C:\IF_EPOC\COM\LOG", 0)
            Case "6":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 6 C:\IF_EPOC\COM\LOG", 0)
            Case "7":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 7 C:\IF_EPOC\COM\LOG", 0)
            Case "8":        Call WinExec("C:\IF_EPOC\COM\EPOC.exe 8 C:\IF_EPOC\COM\LOG", 0)
        End Select
        FileEPOC.PATH = "C:\IF_EPOC\COM\LOG"
    End If
    
    If gHOSP.MACHNM = "UROMETER" Then
        fraUrometer.Visible = True
        txtUro.Text = ""
        cboUro.Clear
        For i = 1 To UBound(gArrEQP)
            cboUro.AddItem gArrEQP(i, 6) & Space(20) & "|" & gArrEQP(i, 6) & "|" & colSTATE + i
        Next
        cboUro.ListIndex = 2
    Else
        fraUrometer.Visible = False
        fraBarcode.LEFT = cmdWorkList.LEFT + cmdWorkList.WIDTH + 100
    End If
    
    
    
    If gDETAILVIEW = "Y" Then
        spdResult.Visible = True
    Else
        spdResult.Visible = False
    End If
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If gComm.COMPORT = "" Then
            If (MsgBox("포트 번호가 지정되지 않았습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
                MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
                
                MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
                MDIIF.imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
                MDIIF.imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
                
                Resume Next
            Else
                End
            End If
        Else
            If (MsgBox("포트 번호(COM" & gComm.COMPORT & ")가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
                MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
                
                MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
                MDIIF.imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
                MDIIF.imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
                
                Resume Next
            Else
                End
            End If
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If

End Sub
 
'
Public Sub OpenCommunication()

    If gComm.COMTYPE = "1" Then

        comEqp.CommPort = gComm.COMPORT
        comEqp.RTSEnable = gComm.RTSEnable
        comEqp.DTREnable = gComm.DTREnable
        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

        If comEqp.PortOpen = False Then
            comEqp.PortOpen = True
        End If

        If comEqp.PortOpen Then
            MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
            MDIIF.imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            MDIIF.imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            MDIIF.imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgOn.ZOrder 0

        Else
            MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            MDIIF.imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            MDIIF.imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgOff.ZOrder 0
        
           ' imgCom.Picture = imlStatus.ListImages("OFF").ExtractIcon
        
        End If
        
    ElseIf gComm.COMTYPE = "2" Then
        'lblComStatus.Left = imgPort.Left + 500
        'lblComStatus.Width = 6000
        If gComm.TCPTYPE = "SERVER" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen
            
            MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 연결중.."

            MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgSend.Visible = False
            'imgReceive.Visible = False
            'lblSend.Visible = False
            'lblRcv.Visible = False
            imgOff.ZOrder 0

        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
            
            MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 연결중..."

            MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgSend.Visible = False
            'imgReceive.Visible = False
            'lblSend.Visible = False
            'lblRcv.Visible = False
            imgOff.ZOrder 0
        
        End If
    ElseIf gComm.COMTYPE = "" Then

    End If

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    MDIIF.cmdNode.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT '- picTop.HEIGHT

     If gWORKPOS = "M" Then
        spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - fraWorkInfo.HEIGHT - 160

        If spdResult.Visible = True Then
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = 140
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100

            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT '/ 2
            
            'imgResult.TOP = spdOrder.TOP + spdResult.HEIGHT
            'imgResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            'imgResult.HEIGHT = spdOrder.HEIGHT / 2
            'imgResult.WIDTH = spdResult.WIDTH
        Else
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = 140
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100
        End If
    Else
        If spdResult.Visible = True Then
            spdOrder.LEFT = 40
            spdOrder.TOP = 140
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200

            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT ' / 2
        
            'imgResult.TOP = spdOrder.TOP + spdResult.HEIGHT
            'imgResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            'imgResult.HEIGHT = spdOrder.HEIGHT / 2
            'imgResult.WIDTH = spdResult.WIDTH
        Else
            spdOrder.LEFT = 40
            spdOrder.TOP = 140
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - 200

            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
        End If
    End If
        
End Sub

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Private Function GetPatTRestResult(ByVal asRow As Integer) As Integer
    Dim strBarno    As String
    Dim intSeq      As String
    Dim strExamDate As String
    Dim intRow   As Integer

On Error GoTo ErrHandle

    GetPatTRestResult = -1
    intRow = 0

    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = GetText(spdOrder, asRow, colEXAMDATE)
    strBarno = GetText(spdOrder, asRow, colBARCODE)
    
    If intSeq = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EQUIPCODE, EXAMNAME, EXAMCODE, EQUIPRESULT, RESULT, PREVRESULT, REFJUDGE, REFFLAG" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE = '" & strBarno & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmInterface.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount - 1
            Do Until AdoRs_Local.EOF
                If AdoRs_Local.Fields("EXAMNAME").Value & "" <> "96M" Then
                    intRow = intRow + 1
                    If AdoRs_Local.Fields("EXAMCODE").Value & "" = "" Then
                        Call SetText(frmInterface.spdResult, "0", intRow, colCHECKBOX)
                    Else
                        Call SetText(frmInterface.spdResult, "1", intRow, colCHECKBOX)
                    End If
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colRSEQNO)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
    '                If AdoRs_Local.Fields("REFJUDGE").Value & "" = "H" Then
    '                    .ForeColor = vbRed
    '                    .FontBold = True
    '                ElseIf AdoRs_Local.Fields("REFJUDGE").Value & "" = "L" Then
    '                    .ForeColor = vbBlue
    '                    .FontBold = True
    '                Else
    '                    .ForeColor = vbBlack
    '                    .FontBold = False
    '                End If
    '                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("PREVRESULT").Value & "", intRow, colRPREVRESULT)
                    Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("REFFLAG").Value & "", intRow, colRFLAG)
                End If
                AdoRs_Local.MoveNext
            Loop
            '.RowHeight(-1) = gROWHEIGHT ' = 15
        End With
        GetPatTRestResult = 1
    End If

    AdoRs_Local.Close

Exit Function

ErrHandle:
    GetPatTRestResult = -1

    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Function

Private Sub imgPort_DblClick()
    
    If gComm.COMTYPE = "1" And comEqp.PortOpen = True Then
        
        If MsgBox("COM Port Close?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
            comEqp.PortOpen = False
        End If
    ElseIf gComm.COMTYPE = "1" And comEqp.PortOpen = False Then
        
        If MsgBox("COM Port Open?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
            comEqp.CommPort = gComm.COMPORT
            comEqp.RTSEnable = gComm.RTSEnable
            comEqp.DTREnable = gComm.DTREnable
            comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
    
            If comEqp.PortOpen = False Then
                comEqp.PortOpen = True
            End If
        End If
    
    End If
    
    If comEqp.PortOpen Then
        MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
        
        MDIIF.imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        MDIIF.imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        MDIIF.imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    
    Else
        MDIIF.lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        
        MDIIF.imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        MDIIF.imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        MDIIF.imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    End If

End Sub



Private Sub lblSlipCd_DblClick()
    Dim strSlipCd   As String
    
    strSlipCd = InputBox("SLIP 코드입력", "SLIP CD", lblSlipCd.Caption)
        
    If strSlipCd <> "" Then
        lblSlipCd.Caption = strSlipCd
    End If

End Sub


Public Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    Dim strItems    As String
    Dim strPName    As String
    Dim strPSex     As String
    Dim strPAge     As String
    
On Error GoTo Err
    '-- 정렬
'    If Row = 0 Then
'        '-- 정렬 추가
'        Exit Sub
'    End If

    If spdOrder.DataRowCnt <= 0 Then
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "1", i, colCHECKBOX)
            Next
        End If
        Exit Sub
    End If
    
'    If Row > 0 And Col = colCHECKBOX Then
'        If GetText(spdOrder, Row, colCHECKBOX) = "1" Then
'            Call SetText(spdOrder, "0", Row, colCHECKBOX)
'        Else
'            Call SetText(spdOrder, "1", Row, colCHECKBOX)
'        End If
'        Exit Sub
'    End If
    
    If Row = 0 Then
        Exit Sub
    End If
    
    '-- 환자정보표시
    lblBarcode.Caption = GetText(spdOrder, Row, colBARCODE)
    
    strPName = GetText(spdOrder, Row, colPNAME)
    strPSex = GetText(spdOrder, Row, colPSEX)
    strPSex = IIf(strPSex = "", "-", strPSex)
    strPAge = GetText(spdOrder, Row, colPAGE)
    strPAge = IIf(strPAge = "", "-", strPAge)
    
    lblPatNm.Caption = strPName & Space(1) & strPSex & "/" & strPAge
    
    lblStatus.Caption = IIf(GetText(spdOrder, Row, colSTATE) = "", "검사준비", GetText(spdOrder, Row, colSTATE))
    
    If chkAdd.Value = "1" Then
        txtOldBarNum.Text = GetText(spdOrder, Row, colBARCODE)
    Else
        txtOldBarNum.Text = ""
    End If
    
    '-- 결과표시
    If GetPatTRestResult(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.MaxRows = 0
        strItems = ""
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◇
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    strItems = strItems & GetText(spdOrder, 0, intCol) & "/"
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = gROWHEIGHT - 5
                End If
            Next
        End With
    End If

    'LoadPicture(App.PATH & "\ICON\" & gHOSP.PARTNM & ".ico")
'    imgResult.Picture = LoadPicture(gB4C.IMAGE & GetText(spdOrder, Row, colEXAMDATE) & "\" & GetText(spdOrder, Row, colBARCODE) & "-" & GetText(spdOrder, Row, colPNAME) & ".jpg")

    lblRow.Caption = Row
    
    If gHOSP.MACHNM = "UROMETER" Then
        txtWBC.Text = ""
        txtRBC.Text = ""
        txtWBC.SetFocus
    End If
    
Exit Sub
Err:
'    If Err.Number = 53 Then
'        MsgBox "결과 이미지가 없습니다.", vbOKOnly + vbCritical, Me.Caption
'    End If
End Sub

Private Sub spdOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim strSEX      As String
    Dim strAGE      As String
    Dim strCREA     As String
    Dim streGFR     As String
    
    
    If Row = 0 Then
        Exit Sub
    End If
    
    
    'frmPreview.Image1.Picture = LoadPicture(gB4C.IMAGE & GetText(spdOrder, Row, colEXAMDATE) & "\" & GetText(spdOrder, Row, colBARCODE) & "-" & GetText(spdOrder, Row, colPNAME) & ".jpg")
    frmPreview.LblPath.Caption = gB4C.IMAGE & GetText(spdOrder, Row, colEXAMDATE) & "\" & GetText(spdOrder, Row, colBARCODE) & "-" & GetText(spdOrder, Row, colPNAME) & ".jpg"
    frmPreview.Image1.Picture = LoadPicture(gB4C.IMAGE & GetText(spdOrder, Row, colEXAMDATE) & "\" & GetText(spdOrder, Row, colBARCODE) & "-" & GetText(spdOrder, Row, colPNAME) & ".jpg")
    frmPreview.Show
    
        
''    'eGFR계산
''    If Mid(GetText(spdOrder, 0, Col), 1, 3) = "CRE" Then
''        strCREA = GetText(spdOrder, Row, Col)
''        If strCREA <> "" And IsNumeric(strCREA) Then
''            '18세 이상만 적용
''            If strAGE <> "" And IsNumeric(strAGE) Then 'And CCur(strAGE) > 18
''                If CCur(strAGE) > 18 Then
''                    streGFR = ""
''
''                    '-- MDRD 공식
''                    'If mPatient.SEX = "M" Then
''                    '    streGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203)
''                    'ElseIf mPatient.SEX = "F" Then
''                    '    streGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203) * 0.742
''                    'End If
''
''                    '--IDMS-MDRD 공식
''                    If strSEX = "M" Then
''                        streGFR = 175 * (strCREA ^ -1.154) * (strAGE ^ -0.203)        'MDRD 공식
''                    Else 'If mPatient.SEX = "F" Then
''                        streGFR = 175 * (strCREA ^ -1.154) * (strAGE ^ -0.203) * 0.742
''                    End If
''
''                    If streGFR <> "" Then
''                        MsgBox streGFR
''                    End If
''                End If
''            End If
''        End If
''    End If
    
End Sub

Private Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    Dim intRow      As Integer
    Dim strSeq      As String
    
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdOrder, sRow, sCol)
    
    If KeyCode = vbKeyReturn Then
        If colBARCODE = sCol Then
            If GetSampleInfo(sRow, spdOrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                    '-- 이미 결과를 받았으면 검사결과를 수정한다.
                    
                    'gPatOrdCd
                    
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdOrder, strSeq + 1, intRow, colSEQNO)
                    strSeq = strSeq + 1
                Next
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If strNewBarNo = "" Then
        
        End If
        
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdOrder, sRow, sRow
        spdOrder.MaxRows = spdOrder.MaxRows - 1
        spdResult.MaxRows = 0
    ElseIf KeyCode = vbKeyDown Then
        DoEvents
        If sRow = spdOrder.MaxRows Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow + 1)
        DoEvents
    ElseIf KeyCode = vbKeyUp Then
        DoEvents
        If sRow = 1 Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow - 1)
        DoEvents
        
    End If
    
End Sub

Private Sub spdOrder_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'    Dim i       As Integer
'
'
'    If y <= 420 Then Exit Sub
'
'
'    With spdOrder
'        If .DataRowCnt <= 0 Then
'            Exit Sub
'        End If
'
'        For i = 1 To .MaxRows
'            Call SetText(spdOrder, "0", i, colCHECKBOX)
'        Next
'
'        For i = .SelBlockRow To .SelBlockRow2
'            Call SetText(spdOrder, "1", i, colCHECKBOX)
'        Next
'    End With
    
End Sub



Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String

    If spdWork.DataRowCnt <= 0 Then
        Exit Sub
    End If

    If Row = 0 Then
        Call SetSpreadSort(spdWork)
        'Exit Sub
    End If
    
    If Row = 0 Then
        
        If Col = colCHECKBOX Then
            If GetText(spdWork, 1, colCHECKBOX) = "1" Then
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "0", i, colCHECKBOX)
                Next
            Else
                For i = 1 To spdWork.DataRowCnt
                    Call SetText(spdWork, "1", i, colCHECKBOX)
                Next
            End If
        Else
            '-- 정렬 추가
            Call SetSpreadSort(spdWork, 0)
        End If
        Exit Sub
    End If

'    If Row > 0 And Col = colCHECKBOX Then
'        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
'            Call SetText(spdWork, "0", Row, colCHECKBOX)
'        Else
'            Call SetText(spdWork, "1", Row, colCHECKBOX)
'        End If
'        Exit Sub
'    End If


End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    Dim strBarno_Work   As String
    'Dim strUritItems    As String
    
    If Row = 0 Then Exit Sub
    
    'If Col <> colBARCODE Then
    '    Exit Sub
    'End If
    
    intWRow = Row
'    spdWork.Row = Row
'
'    spdWork.Col = colBARCODE
'    'spdWork.Col = colCHARTNO
'
'    strBarno_Work = Trim(spdWork.Text)
    
    strBarno_Work = GetText(spdWork, Row, colBARCODE)
    
    With spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            '.Col = colCHARTNO
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            .MaxRows = .MaxRows + 1
            intRow = .MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
            Next
            
            varItems = GetText(spdWork, intWRow, colITEMS)
            varItems = Split(varItems, "/")
            For intItems = 0 To UBound(varItems)
                For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                    .Row = 0
                    .Col = intOCol
                    If varItems(intItems) = Trim(.Text) Then
                        Call SetSPDOrder(frmInterface.spdOrder, intRow, intRow, intOCol, intOCol)
                        Exit For
                    End If
                Next
            Next
            
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
            .RowHeight(-1) = gROWHEIGHT ' = 15
        End If
    
    End With
End Sub

Private Sub spdWork_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strPatName      As String
    Dim sRow            As Long
    
    sRow = spdWork.ActiveRow
    strPatName = Trim(GetText(spdWork, sRow, colPNAME))
    
    If strPatName = "" Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyDelete Then
        If MsgBox(strPatName & " 를 지우시겠습니까?", vbCritical + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        '정보수정
        SQL = ""
        SQL = SQL & "DELETE FROM UB_PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdWork, sRow, colBARCODE)) & "'" & vbCr
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- 성공
            DeleteRow spdWork, sRow, sRow
            spdWork.MaxRows = spdWork.MaxRows - 1
        End If
    End If
End Sub



Private Sub spdWork_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim i       As Integer
    
    If y <= 420 Then Exit Sub
    
    With spdWork
        If .DataRowCnt <= 0 Then
            Exit Sub
        End If
        
        For i = 1 To .MaxRows
            Call SetText(spdWork, "0", i, colCHECKBOX)
        Next
        
        For i = .SelBlockRow To .SelBlockRow2
            Call SetText(spdWork, "1", i, colCHECKBOX)
        Next
    End With

End Sub

Private Sub spdWork_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    Call spdPopUpDel(spdWork, Col, Row, 12)

End Sub

Private Sub tmrBIOLYTE_Timer()
    
    tmrBIOLYTE.Enabled = False
    
    intPhase = 1
    intBufCnt = 0
    
    Call SerialRcvData_BIOLYTE
    
    RcvBuffer = ""
    Erase strRecvData
    
End Sub

Private Sub tmrConn_Timer()
    Dim sqlRet          As Long
    Dim RS          As ADODB.Recordset
    
On Error GoTo ErrHandle
    If gDBTYPE = "1" Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("Select sysdate From DUAL", sqlRet)
        RS.Close
        
        ''Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmInterface.lstComStatus)
    End If
    
    If gDBTYPE = "2" Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("Select sysdate From DUAL", sqlRet)
        RS.Close
        ''Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmInterface.lstComStatus)
    End If
    
    If gDBTYPE = "3" Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("select now()", sqlRet)
        RS.Close
        ''Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmInterface.lstComStatus)
    End If


Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    MDIIF.lblDBStatus.Caption = "데이터베이스 연결실패"
'    frmErrMsg.Show
    
End Sub

Private Sub tmrDBConn_Timer()

    DoEvents

    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
'    MDIIF.lblTestDate.ForeColor = RGB(255, Int((255 * Rnd) + 1), 0)
    
    If MDIIF.imgNet2.Visible = True Then
        MDIIF.imgNet2.Visible = False
        MDIIF.imgNet3.Visible = True
        MDIIF.imgNet3.ZOrder
    Else
        MDIIF.imgNet3.Visible = False
        MDIIF.imgNet2.Visible = True
        MDIIF.imgNet2.ZOrder
    End If
    
End Sub

Private Sub tmrEPOC_Timer()
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim intCnt      As Integer
    Dim strBuf      As String

'On Error GoTo ErrRoutine

    FileEPOC.Refresh
    
    DoEvents
    
    For intIdx = 0 To FileEPOC.ListCount - 1
        FileEPOC.ListIndex = intIdx
        If Right(FileEPOC.PATH, 1) = "\" Then
            strSrcfile = FileEPOC.PATH & FileEPOC.Filename     ' 원본 파일 이름을 정의합니다.
        Else
            strSrcfile = FileEPOC.PATH & "\" & FileEPOC.Filename    ' 원본 파일 이름을 정의합니다.
        End If
        
        Open strSrcfile For Input As #3
        Do While Not EOF(3)
            strBuf = strBuf & Input(1, #3)
        Loop
        Close #3
        
        pBuffer = strBuf
        SetRawData "" & pBuffer
        
        Kill strSrcfile
        
        FileEPOC.Refresh
        
        Call ReceiveProcess

    Next
    
End Sub

Private Sub tmrQ_Timer()
    
    tmrQ.Enabled = False
    If strQState = "Q" Then
        Erase strRecvData
        intSndPhase = 1
        intFrameNo = 1
        comEqp.Output = ENQ
        SetRawData "[Tx]" & ENQ
        
    End If
    
End Sub

Private Sub tmrReceive_Timer()
    
    MDIIF.imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    MDIIF.imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub


Private Sub txtBarNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow As Integer
    
    If KeyCode = vbKeyReturn Then
        If chkAdd.Value = "1" Then
            With spdOrder
                sRow = lblRow.Caption
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                
                Call spdOrder_KeyDown(13, 1)
                
                If GetSampleInfo(sRow, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                Else
                    '정보수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                    SQL = SQL & " ,HOSPDATE = '" & Trim(GetText(spdOrder, sRow, colHOSPDATE)) & "'" & vbCrLf
                    SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                    SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                    SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                    SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                    SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                    SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                    SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                    SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                    SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                    SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                    SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                    SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                End If
                'txtBarNum.Text = ""
                'txtOldBarNum.Text = ""
            End With
        Else
            With spdOrder
                .MaxRows = .MaxRows + 1
                sRow = .MaxRows
                .Row = sRow
                .Col = colBARCODE
                .Text = txtBarNum.Text
                
                If GetSampleInfo(.Row, spdOrder) = -1 Then
                    MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                Else
                    '정보수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                    SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                    SQL = SQL & " ,HOSPDATE = '" & Trim(GetText(spdOrder, sRow, colHOSPDATE)) & "'" & vbCrLf
                    SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                    SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                    SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                    SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                    SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                    SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                    SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                    SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                    SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                    SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                    SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                    SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                    lblRow.Caption = lblRow.Caption + 1
                End If
                
                Call spdActiveCell(spdOrder, .Row + 1, colBARCODE)
                
            End With
        End If
        
        txtBarNum.Text = ""
        
        txtBarNum.SelStart = 0
        txtBarNum.SelLength = Len(txtBarNum.Text)
    
    End If

End Sub





Private Sub txtPosNo_KeyPress(KeyAscii As Integer)
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intRow      As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intRackNo = txtRackNo.Text
        intPosNo = txtPosNo.Text
        
        If Not IsNumeric(intPosNo) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intRackNo, intRow, colRACKNO)
                Call SetText(spdOrder, ((intPosNo Mod 11) + 1) - 1, intRow, colPOSNO)
                intPosNo = intPosNo + 1
                If (intPosNo Mod 11) = 0 Then
                    intRackNo = intRackNo + 1
                    intPosNo = 1
                End If
            Next
        End With
        
        txtRackNo.Text = intRackNo
        txtPosNo.Text = intPosNo
        
        'Call txtSeqNo_KeyPress(vbKeyReturn)
        
    End If
End Sub

Private Sub txtRackNo_KeyPress(KeyAscii As Integer)
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intRow      As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intRackNo = txtRackNo.Text
        intPosNo = txtPosNo.Text
        
        If Not IsNumeric(intRackNo) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            If .MaxRows = 0 Then
                Exit Sub
            End If
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intRackNo, intRow, colRACKNO)
                Call SetText(spdOrder, ((intPosNo Mod 11) + 1) - 1, intRow, colPOSNO)
                intPosNo = intPosNo + 1
                If (intPosNo Mod 11) = 0 Then
                    intRackNo = intRackNo + 1
                    intPosNo = 1
                End If
            Next
        End With
        
        txtRackNo.Text = intRackNo
        txtPosNo.Text = intPosNo
    
        'Call txtSeqNo_KeyPress(vbKeyReturn)
    
    End If
    
'    intRackNo = txtRackNo.Text
'    intPosNo = txtPosNo.Text
'    intSeq = txtSeqNo.Text
'
'    With spdWork
'        For i = 1 To .MaxRows
'            Call SetText(spdWork, Format(intRackNo, "0"), i, colRACKNO)
'            Call SetText(spdWork, ((intPosNo Mod 11) + 1) - 1, i, colPOSNO)
'            Call SetText(spdWork, intSeq, i, colSEQNO)
'            intSeq = intSeq + 1
'            intPosNo = intPosNo + 1
'            If (intPosNo Mod 11) = 0 Then
'                intRackNo = intRackNo + 1
'                intPosNo = 1
'            End If
'
'            txtRackNo.Text = intRackNo
'            txtPosNo.Text = intPosNo
'            txtSeqNo.Text = intSeq
'        Next
'    End With
    
End Sub

Private Sub txtRBC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If lblBarcode.Caption <> "" And lblRow.Caption <> "" Then
            gRow = lblRow.Caption
            Call ResultProcess(lblBarcode.Caption, "RBC", txtRBC.Text, txtRBC.Text)
            txtWBC.SetFocus
        End If
    End If

End Sub

Private Sub txtSeqNo_KeyPress(KeyAscii As Integer)
    Dim intSeq  As Integer
    Dim intRow  As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intSeq = txtSeqNo.Text
        
        If Not IsNumeric(intSeq) Then
            MsgBox "숫자만 입력이 가능합니다"
            Exit Sub
        End If
        
        With spdOrder
            For intRow = .ActiveRow To .MaxRows
                Call SetText(spdOrder, intSeq, intRow, colSEQNO)
                intSeq = intSeq + 1
            Next
        End With
        
        txtSeqNo.Text = intSeq
        
        'Call txtRackNo_KeyPress(vbKeyReturn)
    End If
    
End Sub

Private Sub txtUro_KeyPress(KeyAscii As Integer)
    
    If txtUro <> "" Then
        If KeyAscii = vbKeyReturn Then
            Call cmdUroSet_Click
        End If
    End If
    
End Sub


Private Sub txtWBC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If lblBarcode.Caption <> "" And lblRow.Caption <> "" Then
            gRow = lblRow.Caption
            Call ResultProcess(lblBarcode.Caption, "WBC", txtWBC.Text, txtWBC.Text)
            txtRBC.SetFocus
        End If
    End If
    
End Sub

Private Sub wSCK_Close()
        
    If gComm.TCPTYPE = "SERVER" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    End If

End Sub

Private Sub wSck_Connect()
        
    MDIIF.imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If gComm.TCPTYPE = "SERVER" Then
        MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    Else
        MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트 연결성공"
        imgOn.ZOrder 0
    End If


End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        MDIIF.imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "SERVER" Then
            MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPPORT & " 포트 연결성공"
            imgOn.ZOrder 0
        Else
            MDIIF.lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트 연결성공"
            imgOn.ZOrder 0
        End If
    End If
            
End Sub

Private Sub wSCK_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    wSck.GetData strText
    
    pBuffer = strText

    SetRawData "[Rx]" & pBuffer
    
    Call ReceiveProcess

End Sub

