VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIIF 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SANSOFT Interface"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20115
   Icon            =   "MDIIF.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   20115
      TabIndex        =   2
      Top             =   0
      Width           =   20115
      Begin HSCotrol.CButton cmdPrint 
         Height          =   495
         Left            =   20730
         TabIndex        =   29
         ToolTipText     =   "선택한 검체결과를 OCS/EMR로 저장합니다."
         Top             =   30
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "화면출력"
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
         Picture         =   "MDIIF.frx":25CA
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":341C
      End
      Begin VB.Frame fraJWINFO 
         BackColor       =   &H00F8E4D8&
         Height          =   495
         Left            =   22410
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   2595
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "외래"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   2
            Left            =   1740
            TabIndex        =   28
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "입원"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   1
            Left            =   930
            TabIndex        =   27
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "전체"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   26
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Left            =   1530
         Top             =   -180
      End
      Begin VB.Frame fraStatus 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   14190
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   8655
         Begin HSCotrol.CButton cmdXML 
            Height          =   405
            Left            =   4380
            TabIndex        =   30
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   714
            BackColor       =   15698777
            Caption         =   "XML 정리"
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
         Begin VB.Label lblIFStatus 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   4380
            TabIndex        =   23
            Top             =   90
            Width           =   3075
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblComStatus 
            Appearance      =   0  '평면
            BackColor       =   &H00F8E4D8&
            Caption         =   "Com1 연결성공"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   2070
            TabIndex        =   22
            Top             =   30
            Width           =   2115
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   2595
            Picture         =   "MDIIF.frx":3CF6
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   3915
            Picture         =   "MDIIF.frx":4280
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lblSend 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "받는신호"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   1785
            TabIndex        =   21
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblRcv 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "보내는신호"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   2940
            TabIndex        =   20
            Top             =   300
            Width           =   900
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   1755
            Picture         =   "MDIIF.frx":480A
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgNet1 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":4D94
            Top             =   210
            Width           =   240
         End
         Begin VB.Image imgNet2 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":4EDE
            Top             =   210
            Width           =   240
         End
         Begin VB.Image imgNet3 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":5028
            Top             =   210
            Width           =   240
         End
         Begin VB.Label lblDBStatus 
            BackStyle       =   0  '투명
            Caption         =   "데이터베이스 연결성공"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   465
            Left            =   390
            TabIndex        =   19
            Top             =   90
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   9165
         TabIndex        =   4
         Top             =   -60
         Width           =   3645
         Begin HSCotrol.CButton cmdTestNmSave 
            Height          =   495
            Left            =   2100
            TabIndex        =   14
            ToolTipText     =   "검사자ID/명을 변경합니다."
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   873
            BackColor       =   16777215
            Caption         =   "사용자변경"
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
            Picture         =   "MDIIF.frx":5172
            MaskColor       =   0
            PicCapAlign     =   2
            BorderStyle     =   1
            BorderColor     =   32768
            HoverColor      =   16711680
         End
         Begin VB.TextBox txtTestID 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   870
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   870
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "검사자명 :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lblTestID 
            BackStyle       =   0  '투명
            Caption         =   "검사자ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   13
            ToolTipText     =   "검사자ID를 변경하려면 더블클릭 하세요"
            Top             =   150
            Width           =   975
         End
         Begin VB.Label lblTestNm 
            BackStyle       =   0  '투명
            Caption         =   "검사자명"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   960
            TabIndex        =   12
            ToolTipText     =   "검사자명을 변경하려면 더블클릭 하세요"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "검사자ID :"
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
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   150
            Width           =   825
         End
      End
      Begin VB.CheckBox chkLock 
         BackColor       =   &H00F8E4D8&
         Caption         =   "메뉴고정"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   705
      End
      Begin HSCotrol.CButton cmdClear 
         Height          =   495
         Left            =   4620
         TabIndex        =   6
         ToolTipText     =   "인터페이스 화면을 지웁니다."
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "화면정리"
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
         Picture         =   "MDIIF.frx":5C97
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":6971
      End
      Begin HSCotrol.CButton cmdSave 
         Height          =   495
         Left            =   7530
         TabIndex        =   7
         ToolTipText     =   "선택한 검체결과를 OCS/EMR로 저장합니다."
         Top             =   60
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "선택저장"
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
         Picture         =   "MDIIF.frx":79F8
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":82D2
      End
      Begin HSCotrol.CButton cmdView 
         Height          =   495
         Left            =   6060
         TabIndex        =   8
         ToolTipText     =   "선택한 검체의 상세결과를 보입니다."
         Top             =   60
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "상세결과"
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
         Picture         =   "MDIIF.frx":8BAC
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":9833
      End
      Begin HSCotrol.CButton cmdOrderSend 
         Height          =   495
         Left            =   12840
         TabIndex        =   31
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         BackColor       =   32768
         Caption         =   "오더전송"
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
         Picture         =   "MDIIF.frx":A885
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin VB.Label lblMenuInfo 
         BackStyle       =   0  '투명
         Caption         =   "UROMETER120"
         BeginProperty Font 
            Name            =   "Segoe UI Historic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1440
         TabIndex        =   24
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "검사일자 : "
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
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   90
         Width           =   975
      End
      Begin VB.Label lblTestDate 
         BackStyle       =   0  '투명
         Caption         =   "1971-03-11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   330
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   810
         Picture         =   "MDIIF.frx":AAE1
         Top             =   90
         Width           =   2580
      End
   End
   Begin VB.PictureBox picNode 
      Align           =   3  '왼쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   8700
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   615
      Width           =   3000
      Begin HSCotrol.CButton cmdNode 
         Height          =   9855
         Left            =   2625
         TabIndex        =   5
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   17383
         BackColor       =   16777215
         Caption         =   "◀"
         ForeColor       =   12553049
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   14445
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   25479
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSubList(1)"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   11
      Left            =   4680
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":C2F0
            Key             =   "LIS1101"
            Object.Tag             =   "Menu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":D342
            Key             =   "LIS1102"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":E394
            Key             =   "LIS1104"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":F3E6
            Key             =   "LIS1103"
            Object.Tag             =   "SubMenu"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   " 파일 "
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu00 
      Caption         =   "  인터페이스 "
      Visible         =   0   'False
      Begin VB.Menu mnuHoriba 
         Caption         =   " HORIBA "
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  조회업무 "
      Visible         =   0   'False
      Begin VB.Menu mnuResult 
         Caption         =   " 결과 조회"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   " 워크 조회"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " 설정업무 "
      Visible         =   0   'False
      Begin VB.Menu mnuComm 
         Caption         =   " 통신 설정"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   " 검사 설정"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   " 화면 설정"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   " 옵션 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHosp 
         Caption         =   " 기관정보 설정"
      End
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMRInfo 
         Caption         =   " 전산정보 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu06 
      Caption         =   " 기능 "
      Begin VB.Menu mnuWorkSave 
         Caption         =   " 워크리스트 저장 "
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWorkOpen 
         Caption         =   " 워크리스트 열기"
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " 옵션 "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "▷ 바코드 사용"
         Begin VB.Menu mnuBarcode 
            Caption         =   "바코드사용"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "순번사용"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "체크순"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "▷ 적용 결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "▷ 결과 전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "▷ EMR 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " 기타 "
      WindowList      =   -1  'True
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "원격지원(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnComStatus 
         Caption         =   "통신상태보기"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
      Begin VB.Menu mnuSep28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "SANIF 정보"
      End
   End
End
Attribute VB_Name = "MDIIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

Private Sub chkLock_Click()
    Dim strMenuLock As String
    
    If chkLock.Value = "1" Then
        strMenuLock = "1"
    Else
        strMenuLock = "0"
    End If
    
    'Call WritePrivateProfileString("HOSP", "MENULOCK", strMenuLock, App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MENULOCK", strMenuLock)

End Sub

Private Sub cmdClear_Click()

    If frmInterface.WindowState = 2 Then
        Call frmInterface.frmClear
    End If
    
End Sub

Private Sub cmdNode_Click()
    
'    Call FrmMove

        With MDIIF
            If .cmdNode.Caption = "▶" Then
                .cmdNode.Caption = "◀"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "▶"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With

End Sub

Private Sub cmdNode_MouseIn()
    
    Call FrmMove

End Sub

Private Sub cmdOrderSend_Click()

    If frmInterface.spdOrder.MaxRows < 1 Then
        MsgBox "검사대상자가 없습니다.", vbOKOnly + vbCritical, Me.Caption
    Else
        intPhase = 3
        intSndPhase = 1
        strState = "Q"
        Call frmInterface.SendData(ENQ)
    End If
    
End Sub

Private Sub cmdPrint_Click()
    
    With frmInterface
        .spdOrder.PrintOrientation = PrintOrientationLandscape       '세로출력
        .spdOrder.Action = 13
    End With
    
End Sub

Private Sub cmdSave_Click()
    Dim intRow      As Integer
    Dim intRes      As Integer
    Dim strRCnt     As String
    Dim intRCnt     As Integer
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    If frmInterface.spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("선택한 결과를 전송하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
        With frmInterface.spdOrder
            For intRow = 1 To .DataRowCnt
                strRCnt = GetText(frmInterface.spdOrder, intRow, colRCNT)
                If Not IsNumeric(strRCnt) Then
                    intRCnt = 0
                Else
                    intRCnt = strRCnt
                End If
                .Row = intRow
                .Col = colCHECKBOX
                If .Value = 1 And GetText(frmInterface.spdOrder, intRow, colSTATE) <> "" And intRCnt > 0 Then
                    intRes = SaveTransData(intRow, frmInterface.spdOrder)
                    Call SetUpdateStatus(frmInterface.spdOrder, intRow, intRes)
                End If
            Next
        End With
    End If
End Sub

Private Sub cmdView_Click()

    With frmInterface
        If .WindowState = 2 Then
            If gWORKPOS = "M" Then
                If .spdResult.Visible = False Then
                    .spdResult.Visible = True
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "Y")
                    
                    .spdOrder.HEIGHT = Me.ScaleHeight - .picHeader.HEIGHT - 100
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdWork.WIDTH - .spdResult.WIDTH - 200
                    
                    .spdResult.LEFT = .spdOrder.LEFT + .spdOrder.WIDTH + 50
                    .spdResult.HEIGHT = .spdOrder.HEIGHT
                    .spdResult.TOP = .spdOrder.TOP
                Else
                    .spdResult.Visible = False
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdWork.WIDTH - 200
                    
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "N")
                End If
            Else
                If .spdResult.Visible = False Then
                    .spdResult.Visible = True
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "Y")
                    
                    .spdOrder.HEIGHT = Me.ScaleHeight - .picHeader.HEIGHT - 100
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdResult.WIDTH - 200
                    
                    .spdResult.LEFT = .spdOrder.LEFT + .spdOrder.WIDTH + 50
                    .spdResult.HEIGHT = .spdOrder.HEIGHT
                    .spdResult.TOP = .spdOrder.TOP
                Else
                    .spdResult.Visible = False
                    .spdOrder.WIDTH = Me.ScaleWidth - 200
                    
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "N")
                End If
            End If
        End If
    End With
End Sub


Private Sub lblMenuInfo_Click()

    frmInterface.ZOrder 0
    
End Sub

Private Sub lblMenuInfo_DblClick()

    If fraStatus.Visible = False Then
        fraStatus.Visible = True
    Else
        fraStatus.Visible = False
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    'MDI폼 크기
    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.TOP = IIf(gForm.TOP < 0, 0, gForm.TOP)
        Me.LEFT = IIf(gForm.LEFT < 0, 0, gForm.LEFT)
        Me.WIDTH = IIf(gForm.WIDTH < 0, 10000, gForm.WIDTH)
        Me.HEIGHT = IIf(gForm.HEIGHT < 0, 10000, gForm.HEIGHT)
    End If
    
    cmdNode.HEIGHT = TreeView1.HEIGHT
    Me.Caption = "SANSOFT 인터페이스"
    lblMenuInfo.Caption = gHOSP.MACHNM '"인터페이스"
    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    lblTestID.Caption = gHOSP.USERID
    lblTestNm.Caption = gHOSP.USERNM
    
    Call SetTreeNode
    Call FrmMove
    Call frmShow(frmInterface)
    chkLock.Value = gHOSP.MENULOCK

    If gEMR = "JWINFO" Then
        fraJWINFO.Visible = True
    Else
        fraJWINFO.Visible = False
    End If
    
    If InStr(gHOSP.MACHNM, "BATCH") > 0 Then
        cmdOrderSend.Visible = True
    Else
        cmdOrderSend.Visible = False
    End If
    
    fraStatus.Visible = True


End Sub

'-----------------------------------------------------------------------------'
'   기능 : 이건 별루...
'-----------------------------------------------------------------------------'
Public Sub FrmMove()
    
    If chkLock.Value = "0" Then
        With MDIIF
            If .cmdNode.Caption = "▶" Then
                .cmdNode.Caption = "◀"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "▶"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With
    End If
End Sub

Private Sub SetTreeNode()
    Dim nodX As Node

    picNode.Visible = True
    
    With TreeView1
        .Refresh
        .Visible = False
        .LabelEdit = lvwManual
        
        .ImageList = imlSubList(11)
        .HideSelection = False
        .Nodes.Clear
        
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS000", "인터페이스", "LIS1101")
        .Nodes("LIS000").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS001", "조회업무", "LIS1101")
        .Nodes("LIS001").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS002", "설정업무", "LIS1101")
        .Nodes("LIS002").Expanded = True
        .LineStyle = tvwTreeLines
        .Indentation = 300
        
        Set nodX = Nothing
        .Visible = True
        
    End With

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Call mnuExit_Click

End Sub

Private Sub MDIForm_Resize()
    
    If Me.WindowState = 2 Then
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN", "Y")
    Else
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN", "N")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "TOP", gForm.TOP)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "LEFT", gForm.LEFT)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "WIDTH", gForm.WIDTH)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "HEIGHT", gForm.HEIGHT)
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnComStatus_Click()
    
    If fraStatus.Visible = True Then
        fraStatus.Visible = False
    Else
        fraStatus.Visible = True
    End If
    
End Sub

Private Sub mnuAbout_Click()

    Call ShowForm(frmAbout, "산소프트 SANIF 정보")

End Sub

Private Sub mnuHoriba_Click()
    
    Call ShowForm(frmInterface, "인터페이스")

End Sub

Private Sub mnuWorkOpen_Click()
    Dim strPath  As String
    Dim TextLine
    Dim strBuffer
    Dim strCount    As String
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    If frmInterface.spdOrder.MaxRows > 0 Then
        If MsgBox("현재 화면을 지우고 워크리스트를 불러오겠습니까?", vbYesNo + vbInformation, "워크리스트 불러오기") = vbNo Then
            Exit Sub
        End If
    End If
    
    With frmInterface.CommonDialog1
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
 
    strCount = mGetP(mGetP(mGetP(strPath, 2, "WL_"), 3, "_"), 1, ".")
    
    frmInterface.spdOrder.MaxRows = strCount
    
    With frmInterface.spdOrder
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .Clip = strBuffer
        .ClipboardPaste
        .BlockMode = False
        
        .RowHeight(-1) = gROWHEIGHT
    End With
    
Exit Sub
ErrHandler:
                        
End Sub

Private Sub mnuWorkSave_Click()
    Dim strBuffer As String
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    With frmInterface
        If .spdOrder.MaxRows < 1 Then
            MsgBox "저장할 워크리스트가 없습니다.", vbOKOnly + vbCritical, "워크 리스트"
            Exit Sub
        End If
        
        Call .spdOrder.SetSelection(1, 1, .spdOrder.MaxCols, .spdOrder.MaxRows)
        '클립보드 카피
        .spdOrder.ClipboardCopy
        
        strBuffer = Clipboard.GetText()
        
        Call SetWorkData(strBuffer, .spdOrder.MaxRows)
        
        MsgBox "워크리스트가 저장 되었습니다.", vbOKOnly + vbInformation, "워크 리스트"
        
    End With
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Call TreeFromLoad(Node)
    
End Sub

Private Sub TreeFromLoad(ByVal Button As MSComctlLib.Node, Optional ByVal intIdx As Integer)
    
    If Button.Children <> 0 Then
        Exit Sub
    End If
    
    With TreeView1
        Select Case Button.Key
            '인터페이스 ===========================================================================================================
            Case "LIS000":
                            TreeView1.Nodes.Add "LIS000", tvwChild, "LIS00001", gHOSP.MACHNM, "LIS1103"
                            'TreeView1.Nodes.Add "LIS000", tvwChild, "LIS00002", "XP300", "LIS1103"
                            
                            Case "LIS00001":        Call ShowForm(frmInterface, frmInterface.Caption)
                            'Case "LIS00002":        Call ShowForm(frmInterface2, frmInterface2.Caption)
                            
            '조회업무 ===========================================================================================================
            Case "LIS001":
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00101", "결과 조회", "LIS1103"
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00102", "워크 조회", "LIS1103"
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00103", "검사 통계", "LIS1103"

                            Case "LIS00101":        Call ShowForm(frmResult, frmResult.Caption)
                            Case "LIS00102":        Call ShowForm(frmWorkList, frmWorkList.Caption)
                            Case "LIS00103":        Call ShowForm(frmStatistics, frmStatistics.Caption)
            '설정업무 =======================================================================================================
            Case "LIS002":
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00201", "검사설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00202", "통신설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00203", "화면설정", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00204", "기관정보설정", "LIS1103"
                            'TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00205", "옵션설정", "LIS1103"

                            Case "LIS00201":        Call ShowForm(frmTestSet, frmTestSet.Caption)
                            Case "LIS00202":        Call ShowForm(frmConfig, frmConfig.Caption)
                            Case "LIS00203":        Call ShowForm(frmScreenSet, frmScreenSet.Caption)
                            Case "LIS00204":        Call ShowForm(frmHospInfo, frmHospInfo.Caption)
                            'Case "LIS00205":        Call ShowForm(frmTestOptSet, frmTestOptSet.Caption)
            
            
            
        End Select
    End With
    
End Sub

Private Sub cmdTestNmSave_Click()
    
    If txtTestID.Text <> "" Then
        lblTestID.Caption = txtTestID.Text
        'Call WritePrivateProfileString("HOSP", "USERID", lblTestID.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERID", txtTestID.Text)
        
        txtTestID.Visible = False
        lblTestID.Visible = True
    End If
    
    If txtTestNm.Text <> "" Then
        lblTestNm.Caption = txtTestNm.Text
        'Call WritePrivateProfileString("HOSP", "USERNM", lblTestNm.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERNM", txtTestNm.Text)
        txtTestNm.Visible = False
        lblTestNm.Visible = True
    End If
    
End Sub


Private Sub lblTestID_DblClick()
    If txtTestID.Visible = False Then
        txtTestID.Text = lblTestID.Caption
        lblTestID.Visible = False
        txtTestID.Visible = True
    Else
        txtTestID.Text = ""
        lblTestID.Visible = True
        txtTestID.Visible = False
    End If
End Sub


Private Sub lblTestNm_DblClick()
    If txtTestNm.Visible = False Then
        txtTestNm.Text = lblTestNm.Caption
        lblTestNm.Visible = False
        txtTestNm.Visible = True
    Else
        txtTestNm.Text = ""
        lblTestNm.Visible = True
        txtTestNm.Visible = False
    End If
End Sub


Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "Y")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "0")


End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "3")

End Sub

Private Sub mnuComm_Click()
    
    frmConfig.Show

End Sub

Private Sub mnuComTest_Click()

End Sub

Private Sub mnuCommTest_Click()

    If frmInterface.picComm.Visible = True Then
        frmInterface.picComm.Visible = False
    Else
        frmInterface.picComm.Visible = True
        frmInterface.picComm.ZOrder 0
    End If
    
End Sub

Private Sub mnuEMRInfo_Click()
    
    If InputBox("비밀번호 입력" & Space(5) & "hint:개발자oyh") = "dev0503" Then
        frmEMRInfo.Show
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    'Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS", "N")

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then

        If KillProcess("EPOC") = False Then
            Call Shell("taskkill.exe /im EPOC.exe", 0)
        End If
        
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHosp_Click()

    frmHospInfo.Show 'vbModal

End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    'Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS", "Y")

End Sub

Private Sub mnuOpt_Click()
    
    frmTestOptSet.Show 'vbModal
    
End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "2")

End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show 'vbModal
    
End Sub

Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    'Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO", "Y")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    'Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO", "N")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "1")
    
End Sub

Private Sub mnuTest_Click()
    
    frmTestSet.Show 'vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show 'vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show 'vbModal

End Sub


