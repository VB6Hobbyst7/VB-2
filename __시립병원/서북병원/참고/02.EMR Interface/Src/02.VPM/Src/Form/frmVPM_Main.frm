VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C9460280-3EED-11D0-A647-00A0C91EF7B9}#1.0#0"; "ImageViewer2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmVPM_Main 
   Caption         =   "Hi Interface EMR(VPM)"
   ClientHeight    =   10110
   ClientLeft      =   6060
   ClientTop       =   1650
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVPM_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   15030
   Begin VB.TextBox txtSerialData 
      Height          =   4095
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   52
      Top             =   5580
      Width           =   9915
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin FPSpread.vaSpread sprEQ_INFO 
      Height          =   1335
      Left            =   1740
      TabIndex        =   47
      Top             =   12300
      Visible         =   0   'False
      Width           =   2415
      _Version        =   393216
      _ExtentX        =   4260
      _ExtentY        =   2355
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
      SpreadDesigner  =   "frmVPM_Main.frx":9F8A
   End
   Begin MSComctlLib.ProgressBar prgPatient 
      Height          =   75
      Left            =   5040
      TabIndex        =   46
      Top             =   1080
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "전송(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13980
      Picture         =   "frmVPM_Main.frx":A1AB
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fraLimageList 
      Caption         =   "[Image List]"
      Height          =   3975
      Left            =   60
      TabIndex        =   40
      Top             =   5700
      Width           =   4875
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   960
         ScaleHeight     =   285
         ScaleWidth      =   3825
         TabIndex        =   48
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton optImage전송여부 
            Caption         =   "전송완료"
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optImage전송여부 
            Caption         =   "미전송"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "Double Click 하면 Image 파일로 전환됩니다."
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.PictureBox picImageList미전송 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   4725
         TabIndex        =   50
         Top             =   660
         Width           =   4755
         Begin VB.CommandButton cmdViewLocal 
            Caption         =   "자료불러오기"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteLocal 
            Caption         =   "삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   10
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmd미전송폴더변경 
            Caption         =   "미전송폴더변경"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   2880
            Width           =   1875
         End
         Begin FPSpread.vaSpread spr미전송 
            Height          =   2775
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   4635
            _Version        =   393216
            _ExtentX        =   8176
            _ExtentY        =   4895
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   20
            OperationMode   =   3
            SpreadDesigner  =   "frmVPM_Main.frx":C7E5
         End
      End
      Begin VB.PictureBox picImageList전송완료 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   4725
         TabIndex        =   51
         Top             =   660
         Width           =   4755
         Begin VB.CommandButton cmdViewFTP 
            Caption         =   "자료불러오기"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteFTP 
            Caption         =   "삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   14
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmd전송완료폴더변경 
            Caption         =   "전송완료폴더변경"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   2880
            Width           =   1875
         End
         Begin FPSpread.vaSpread spr전송완료 
            Height          =   2775
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   4635
            _Version        =   393216
            _ExtentX        =   8176
            _ExtentY        =   4895
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   20
            OperationMode   =   3
            SpreadDesigner  =   "frmVPM_Main.frx":CC9D
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전송여부"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[수진자내역]"
      Height          =   4995
      Left            =   60
      TabIndex        =   36
      Top             =   600
      Width           =   4875
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   960
         ScaleHeight     =   285
         ScaleWidth      =   3825
         TabIndex        =   39
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton opt전송여부 
            Caption         =   "미전송"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   0
            ToolTipText     =   "Double Click 하면 Image 파일로 전환됩니다."
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt전송여부 
            Caption         =   "전송완료"
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   1
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "조회(&V)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   600
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker dtp접수일자 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40449
      End
      Begin FPSpread.vaSpread spr수진자내역 
         Height          =   3975
         Left            =   60
         TabIndex        =   4
         Top             =   960
         Width           =   4755
         _Version        =   393216
         _ExtentX        =   8387
         _ExtentY        =   7011
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   20
         OperationMode   =   3
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmVPM_Main.frx":D171
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "전송여부"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   38
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수일자"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar staCondition 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   9735
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5398
            MinWidth        =   3528
            Text            =   "Copyright ⓒ 2010 Medimate Corp."
            TextSave        =   "Copyright ⓒ 2010 Medimate Corp."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12462
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "DB"
            TextSave        =   "DB"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "FTP"
            TextSave        =   "FTP"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "COM"
            TextSave        =   "COM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "PRN"
            TextSave        =   "PRN"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2011-12-13"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "오후 4:55"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   8475
      Left            =   5040
      ScaleHeight     =   8445
      ScaleWidth      =   9885
      TabIndex        =   45
      Top             =   1200
      Width           =   9915
      Begin SCRIBBLELib.ImageViewer imvResult 
         Height          =   7335
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
         _ExtentY        =   12938
         _StockProps     =   0
         LicenseKey      =   "12595"
      End
      Begin VB.PictureBox picControl 
         Height          =   555
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   9735
         TabIndex        =   65
         Top             =   480
         Width           =   9795
         Begin VB.CommandButton cmdCenter 
            Caption         =   "Center Image"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8700
            TabIndex        =   74
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdZmHeight 
            Caption         =   "Zoom to Height"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7620
            TabIndex        =   73
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdZmWidth 
            Caption         =   "Zoom to Width"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6540
            TabIndex        =   72
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdFit 
            Caption         =   "Fit to Window"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5460
            TabIndex        =   71
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdRotate 
            Caption         =   "Rotate"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4380
            TabIndex        =   70
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmd100 
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   69
            Top             =   0
            Width           =   1035
         End
         Begin VB.ComboBox cboZoomValue 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            Style           =   2  '드롭다운 목록
            TabIndex        =   68
            Top             =   60
            Width           =   1035
         End
         Begin VB.CommandButton cmdzoomin 
            Caption         =   "╋"
            Height          =   495
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdzoomout 
            BackColor       =   &H00FFFFFF&
            Caption         =   "━"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   66
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "배율"
            Height          =   180
            Index           =   12
            Left            =   1440
            TabIndex        =   75
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.PictureBox picJPG 
         Appearance      =   0  '평면
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   9765
         TabIndex        =   63
         Top             =   60
         Width           =   9795
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "JPG File"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   64
            Top             =   60
            Width           =   1080
         End
      End
      Begin VB.PictureBox picTifPdf 
         Height          =   435
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   9735
         TabIndex        =   53
         Top             =   60
         Width           =   9795
         Begin VB.CommandButton cmdMultiLast 
            Caption         =   "->|"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   59
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiNext 
            Caption         =   "->"
            Height          =   375
            Left            =   960
            TabIndex        =   58
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiPrev 
            Caption         =   "<-"
            Height          =   375
            Left            =   480
            TabIndex        =   57
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiFirst 
            Caption         =   "|<-"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiJump 
            Caption         =   "Go"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3420
            TabIndex        =   55
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtMultiPno 
            Alignment       =   1  '오른쪽 맞춤
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2940
            TabIndex        =   54
            Text            =   "1"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Total Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4230
            TabIndex        =   62
            Top             =   60
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Page No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   61
            Top             =   60
            Width           =   855
         End
         Begin VB.Label lblMultiCnt 
            Alignment       =   1  '오른쪽 맞춤
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5400
            TabIndex        =   60
            Top             =   60
            Width           =   615
         End
      End
   End
   Begin VB.Label lbl처방상태 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "외래"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10380
      TabIndex        =   77
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방상태"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   15
      Left            =   9420
      TabIndex        =   76
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lbl처방코드 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   44
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   6
      Left            =   7380
      TabIndex        =   43
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방SEQ"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   5
      Left            =   11820
      TabIndex        =   42
      Top             =   600
      Width           =   705
   End
   Begin VB.Label lbl처방SEQ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   41
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비명"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   35
      Top             =   60
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병록번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   5100
      TabIndex        =   34
      Top             =   360
      Width           =   780
   End
   Begin VB.Label lbl병록번호 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   33
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "수진자명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   7380
      TabIndex        =   32
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "연령/성별"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   9420
      TabIndex        =   31
      Top             =   360
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "진료과"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   5100
      TabIndex        =   30
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "환자구분"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   4
      Left            =   7380
      TabIndex        =   29
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lbl수진자명 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "홍길동"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   28
      Top             =   360
      Width           =   585
   End
   Begin VB.Label lbl연령성별 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "40/남"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10380
      TabIndex        =   27
      Top             =   360
      Width           =   510
   End
   Begin VB.Label lbl진료과 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "IM"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   26
      Top             =   840
      Width           =   210
   End
   Begin VB.Label lbl입외구분 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "외래"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   25
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Patient/Image Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5160
      TabIndex        =   24
      Top             =   60
      Width           =   2640
   End
   Begin VB.Label lbl결과일자 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   23
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label lbl처방일자 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   22
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label lbl처방명 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   21
      Top             =   600
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "결과일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   11
      Left            =   11820
      TabIndex        =   20
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방코드"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   8
      Left            =   5100
      TabIndex        =   19
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   7
      Left            =   11820
      TabIndex        =   18
      Top             =   360
      Width           =   780
   End
   Begin VB.Shape shpPatientInfo 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Left            =   5040
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   8835
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4875
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File    "
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "구성    "
      Begin VB.Menu mnuSettingSub 
         Caption         =   "DataBase Info Setting"
         Index           =   0
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "Target Equipment Setting"
         Index           =   1
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "Equipment Config"
         Index           =   2
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "정보"
   End
End
Attribute VB_Name = "frmVPM_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height의 초기값
Dim lngMeWidth      As Long '/Me.Width의 초기값

Private Type ConWhere   ' 사용자 정의 형식을 만듭니다.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Private iPicCnt     As Integer
Private MMFTP       As New cls공용_FTP
Private MMSFTP      As New cls공용_SFTP

Private mintCurFrame As Integer ' 현재 프레임이 보입니다.

Public Function FUNC_MM_CANCEL() As Boolean
    lbl장비명 = ""
    prgPatient.Max = 100
    prgPatient.Value = 100
    
    '/수진자내역
    opt전송여부(0).Value = True
    opt전송여부(0).ForeColor = RGB(0, 0, 255)
    opt전송여부(0).FontBold = True
    
    dtp접수일자.Value = Format(Now, "YYYY-MM-DD")
    
    optImage전송여부(0).ForeColor = RGB(0, 0, 255)
    optImage전송여부(0).FontBold = True
    
    Call FUNC_MM_KEY_CLEAR("1") '/수진자내역 Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    Call FUNC_MM_KEY_CLEAR("4") '/미전송 Spread Clear
    Call FUNC_MM_KEY_CLEAR("5") '/전송완료 Spread Clear
    
    picImageList미전송.Visible = True
    picImageList전송완료.Visible = False

    mnuSetting.Visible = False '/구성메뉴 안보이기

    txtSerialData.Visible = False
    
    picTifPdf.Visible = False
    picJPG.Visible = True
End Function

Public Function FUNC_MM_DELETE(ArgSection As String) As Boolean
    Dim strIMGFILEPATH  As String
    
    FUNC_MM_DELETE = False
    
On Error GoTo ERR_RTN
    
    Select Case ArgSection
        Case "1": GoSub DELETE1_RTN '/미전송 삭제
        Case "2": GoSub DELETE2_RTN '/전송완료 삭제
    End Select
    
    FUNC_MM_DELETE = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

DELETE1_RTN:
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")

    For intX = 1 To spr미전송.MaxRows
        If GET_CELL(spr미전송, 1, intX) = "1" Then
            Kill gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, intX)
        End If
    Next intX
Return

'/----------------------------------------------------------------------------------------------------/

DELETE2_RTN:
    '/----------------------------------------------------------------------------------------------------/
    '/Step4.    FTP 연결
    '/----------------------------------------------------------------------------------------------------/
    If OpenDB(gstrREG_DB_CONSTR) = False Then End
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl병록번호 & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbl처방일자, "-", "") & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbl처방SEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
    
    If Not ADR Is Nothing Then
        strIMGFILEPATH = Trim(ADR!IMGFILEPATH & "")                              '/FTP경로
        
        ADR.Close: Set ADR = Nothing
    End If

    Call CloseDB
    
    Dim success As Long
    success = sftp.IsConnected
    If (success <> 1) Then
        If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
            MsgBox "Image File Server에 접근할 수 없습니다." & vbCrLf & "전산실에 문의바랍니다.", vbCritical, "전송불가"
            Exit Function
        End If
    End If
    
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH) = True Then
        For intX = 1 To spr전송완료.MaxRows
            If GET_CELL(spr전송완료, 1, intX) = "1" Then
'''                If MMFTP.DeleteFTPFile(GET_CELL(spr전송완료, 2, intX)) = True Then
'''                    Kill gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, intX)
'''                Else
'''                    Exit Function
'''                End If
            
                '/서북병원 통합View OCX(광화일 네파인더와 공동제작)파일에서는 JPG 파일만 가져온다. BAK파일은 DownLoad 하지 않으므로 Rename해도 상관없다.
                If MMSFTP.RenameFTPFile(strIMGFILEPATH & GET_CELL(spr전송완료, 2, intX), strIMGFILEPATH & Left(GET_CELL(spr전송완료, 2, intX), InStr(GET_CELL(spr전송완료, 2, intX), ".")) & "bak") = True Then
                    Kill gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, intX)
                Else
                    Exit Function
                End If
            End If
        Next intX
        
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
    
        If UBound(FtpScanFileName_IMG) = 0 Then '/정상적인 FTP자료가 없으면...
            If OpenDB(gstrREG_DB_CONSTR) = False Then End
            
            ADC.BeginTrans
            
            gstrQuy = "DELETE FROM MM_EMR_RES "
            gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl병록번호 & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbl처방일자, "-", "") & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbl처방SEQ) & " "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            '/원처방:   TMPRSCINFN
            gstrQuy = "UPDATE TMPRSCINFN SET "
            gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '440', "                                              '/440.접수
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "                          '/최종수정자 ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSDATE "                                             '/최종수정일자
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(lbl처방일자), "YYYYMMDD") & "' "    '/처방일자
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & lbl처방SEQ & " "                                 '/처방번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & lbl병록번호 & "' "                               '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & lbl처방코드 & "' "                               '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                 '/원처방 살아있는 처방
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                 '/처방History 번호
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

            '/실시처방: TMPRSCEXCN
            gstrQuy = "UPDATE TMPRSCEXCN SET "
            gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '440', "                                              '/440.접수
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "                          '/최종수정자 ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSDATE "                                             '/최종수정일자
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(lbl처방일자), "YYYYMMDD") & "' "    '/처방일자
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & lbl처방SEQ & " "                                 '/처방번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & lbl병록번호 & "' "                               '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & lbl처방코드 & "' "                               '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/실시처방 살아있는 처방
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

            ADC.CommitTrans
            
            Call CloseDB
        End If
        
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
        If UBound(FtpScanFileName) = 0 Then '/BackUp 된 FTP자료도 없으면...
            Call MMSFTP.RemoveFTPDirectory(strIMGFILEPATH)
        End If
    End If
    
    Call MMSFTP.CloseConnection
Return

'/----------------------------------------------------------------------------------------------------/

ERR_RTN:
    MsgBox "삭제 오류!!!", vbCritical, "확인"
End Function

Public Sub FUNC_MM_INITIAL()
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 10890
    lngMeWidth = 15150
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Show
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    Call FUNC_MM_CANCEL
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    
    lbl장비명 = gtypEQ_INFO.EQUIPNM
    
    '/작업 상태 Check----------------------------------------------------------------------------------------------------/
    staCondition.Panels.Item(3).Enabled = False
    staCondition.Panels.Item(4).Enabled = False
    staCondition.Panels.Item(5).Enabled = False
    staCondition.Panels.Item(6).Enabled = False

    staCondition.Panels.Item(5).Visible = False
    staCondition.Panels.Item(6).Visible = False

    If gstrSTAUS_DB = "Y" Then
        staCondition.Panels.Item(3).Enabled = True
    End If
    If gstrSTAUS_FTP = "Y" Then
        staCondition.Panels.Item(4).Enabled = True
    End If
    
    If gtypEQ_INFO.SERIALYN = "Y" Then
        staCondition.Panels.Item(5).Visible = True
        
        On Error GoTo RTN_ERR_PORT
        
        MSComm1.CommPort = gtypEQ_INFO.SERIALPORT
        MSComm1.RTSEnable = gtypEQ_INFO.SERIALRTS
        MSComm1.DTREnable = gtypEQ_INFO.SERIALDTR
        MSComm1.Settings = gtypEQ_INFO.SERIALBAUD & "," & gtypEQ_INFO.SERIALPARITY & "," & gtypEQ_INFO.SERIALDATABIT & "," & gtypEQ_INFO.SERIALSTOPBIT
        
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        
            staCondition.Panels.Item(5).Enabled = True
        End If
    End If

    If gtypEQ_INFO.ZIPYN = "Y" Then
        staCondition.Panels.Item(6).Visible = True
        
        If Trim(gtypEQ_INFO.ZIPNM) <> "" Then
            Dim X As Printer
            
            strTemp = ""
            For Each X In Printers
                If Trim(gtypEQ_INFO.ZIPNM) = X.DeviceName Then
                    staCondition.Panels.Item(6).Enabled = True
                    
                    strTemp = "Y"
                    Exit For
                End If
            Next X
            
            If strTemp = "" Then
                MsgBox "해당 장비는 가상프린터를 사용해야하는 장비입니다." & vbCrLf & _
                       "가상프린터 정보가 없거나 다르므로 가상프린터를 (재)지정하십시오!", vbInformation, "알림"
                frm공용_Set_Equip_Config.Show vbModal
            End If
        Else
            MsgBox "해당 장비는 가상프린터를 사용해야하는 장비입니다." & vbCrLf & _
                   "가상프린터 정보가 없거나 다르므로 가상프린터를 (재)지정하십시오!", vbInformation, "알림"
            frm공용_Set_Equip_Config.Show vbModal
        End If
    End If
    '/작업 상태 Check----------------------------------------------------------------------------------------------------/
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    Me.Caption = Me.Caption & Space(10) & "(사용자: " & gtypUSER.USERNM & " )"
    
    '/Zoom 비율
    cboZoomValue.Clear
    cboZoomValue.AddItem ""
    cboZoomValue.AddItem "25%"
    cboZoomValue.AddItem "33%"
    cboZoomValue.AddItem "50%"
    cboZoomValue.AddItem "75%"
    cboZoomValue.AddItem "100%"
    cboZoomValue.AddItem "150%"
    cboZoomValue.AddItem "200%"
    
    '''Call SUB_GET_REG_CLIENT_INFO '/설정 장비 Image 폴더 변경 후 레지스터 조정할려고 만듦(현재는 의미없음)
Return

'/----------------------------------------------------------------------------------------------------/

RTN_ERR_PORT:
    If Err = 8002 Then      'Port
        staCondition.Panels.Item(5).Enabled = False
        
        MsgBox "통신 포트를 확인하세요!", vbInformation, "알림"
        frm공용_Set_Equip_Config.Show vbModal
    Else
        Resume Next
    End If
End Sub

Public Sub FUNC_MM_KEY_CLEAR(ArgSection As String)
    Select Case ArgSection
        Case "1" '/수진자내역 Spread Clear
            If spr수진자내역.MaxRows > 0 Then spr수진자내역.MaxRows = 0
            
        Case "2" '/Patient Information Clear
            lbl병록번호 = ""
            lbl수진자명 = ""
            lbl연령성별 = ""
            lbl처방코드 = ""
            lbl처방명 = ""
            lbl진료과 = ""
            lbl입외구분 = ""
            lbl처방일자 = ""
            lbl처방SEQ = ""
            lbl결과일자 = ""
            lbl처방상태 = ""
            lbl처방상태.ForeColor = RGB(0, 0, 0)
            
        Case "3": '/Image Clear
            imvResult.Filename = ""
            Call SUB_CHK_PDF_TIF("")
            
        Case "4": '/미전송 Spread Clear
            If spr미전송.MaxRows > 0 Then spr미전송.MaxRows = 0
    
        Case "5": '/전송완료 Spread Clear
            If spr전송완료.MaxRows > 0 Then spr전송완료.MaxRows = 0
    End Select
End Sub

Public Function FUNC_MM_PRINT() As Boolean
'''    Dim strFont1  As String
'''    Dim strFont2  As String
'''    Dim strHead1  As String
'''
'''    If sprVIEW.MaxRows = 0 Then MsgBox "출력할 자료가 없습니다.", vbInformation, "확인": Exit Function
'''
'''    If MsgBox("출력하겠습니까?", vbQuestion + vbOKCancel, "출력여부") = vbCancel Then Exit Function
'''
'''    strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
'''    strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"
'''
'''    strHead1 = "/f1/c" & "거래처 코드" & "/n/n/n"
'''
'''    With sprVIEW
'''        .PrintAbortMsg = "거래처 코드 출력 중..."
'''        .PrintHeader = strFont1 + strHead1 + strFont2
'''        .PrintFooter = "/c" & "PAGE : " & "/P"
'''        .PrintBorder = True
'''        .PrintGrid = True
'''        .PrintColHeaders = True
'''        .PrintRowHeaders = True
'''        .PrintColor = False
'''        .PrintMarginTop = 500
'''        .PrintMarginBottom = 500
'''        .PrintMarginLeft = 500
'''        .PrintMarginRight = 0
'''        .PrintType = PrintTypeAll
'''        .PrintShadows = False
'''        .PrintUseDataMax = False
'''        .Action = ActionSmartPrint
'''    End With
End Function

Public Function FUNC_MM_SAVE(argintRow수진자내역 As Integer) As Boolean
    Dim rtn
    Dim strFileName     As String
    Dim nImageSeq       As Integer
    Dim strTemp1
    Dim strIMGFILEPATH  As String
    
    Dim fstr처방일자    As String
    Dim fstr병록번호    As String
    Dim fstr처방SEQ     As String
    Dim fstr진료과      As String
    Dim fstr처방코드    As String
    Dim fstr입외구분    As String
    
    fstr처방일자 = GET_CELL(spr수진자내역, 7, argintRow수진자내역)
    fstr병록번호 = GET_CELL(spr수진자내역, 2, argintRow수진자내역)
    fstr처방SEQ = GET_CELL(spr수진자내역, 8, argintRow수진자내역)
    fstr진료과 = GET_CELL(spr수진자내역, 4, argintRow수진자내역)
    fstr처방코드 = GET_CELL(spr수진자내역, 11, argintRow수진자내역)
    fstr입외구분 = GET_CELL(spr수진자내역, 3, argintRow수진자내역)
    
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    '/----------------------------------------------------------------------------------------------------/
    '/Step4.    FTP 연결
    '/----------------------------------------------------------------------------------------------------/
    Dim success As Long
    success = sftp.IsConnected
    If (success <> 1) Then
        If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
            MsgBox "Image File Server에 접근할 수 없습니다." & vbCrLf & "전산실에 문의바랍니다.", vbCritical, "전송불가"
            Exit Function
        End If
    End If
   '/----------------------------------------------------------------------------------------------------/
    '/Step5.    FTP 서버의 해당 폴더 생성/ 있으면 SKIP
    '/----------------------------------------------------------------------------------------------------/
    '/기본 Image 폴더 이동
    strIMGFILEPATH = ""
    
    If MMSFTP.SetFTPDirectory("upload/lis") = False Then
        'MsgBox "FTP 서버에 EMR_Image 폴더가 없습니다.", vbInformation, "확인"
        'Exit Sub
        Call MMSFTP.CreateFTPDirectory("upload/lis")
    End If
    strIMGFILEPATH = "upload/lis/"
    
    '/장비코드 폴더 이동 및 생성
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ) = False Then
            MsgBox "FTP 폴더 생성(장비코드&장비SEQ) 중 문제가 발생하였습니다.", vbCritical, "전송불가"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ)
    End If
    strIMGFILEPATH = strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ & "/"

    '/처방년도 폴더 이동 및 생성
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & Left(Replace(fstr처방일자, "-", ""), 4)) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & Left(Replace(fstr처방일자, "-", ""), 4)) = False Then
            MsgBox "FTP 폴더 생성(처방년도) 중 문제가 발생하였습니다.", vbCritical, "전송불가"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & Left(Replace(fstr처방일자, "-", ""), 4))
    End If
    strIMGFILEPATH = strIMGFILEPATH & Left(Replace(fstr처방일자, "-", ""), 4) & "/"
    
    '/처방월일 폴더 이동 및 생성
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & Mid(Replace(fstr처방일자, "-", ""), 5)) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & Mid(Replace(fstr처방일자, "-", ""), 5)) = False Then
            MsgBox "FTP 폴더 생성(처방월일) 중 문제가 발생하였습니다.", vbCritical, "전송불가"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & Mid(Replace(fstr처방일자, "-", ""), 5))
    End If
    strIMGFILEPATH = strIMGFILEPATH & Mid(Replace(fstr처방일자, "-", ""), 5) & "/"

    '/병록번호 폴더 이동 및 생성
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr병록번호) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & fstr병록번호) = False Then
            MsgBox "FTP 폴더 생성(병록번호) 중 문제가 발생하였습니다.", vbCritical, "전송불가"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr병록번호)
    End If
    strIMGFILEPATH = strIMGFILEPATH & fstr병록번호 & "/"

    '/처방SEQ(검진 접수번호) 폴더 이동 및 생성
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr처방SEQ) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & fstr처방SEQ) = False Then
            MsgBox "FTP 폴더 생성(처방SEQ) 중 문제가 발생하였습니다.", vbCritical, "전송불가"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr처방SEQ)
    End If
    strIMGFILEPATH = strIMGFILEPATH & fstr처방SEQ & "/"
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step6.    전송할 확장자 제외한 파일명 정의
    '/----------------------------------------------------------------------------------------------------/
    strFileName = Replace(fstr처방일자, "-", "") & "@" & fstr처방SEQ & "@" & fstr병록번호 & "@"

    '/----------------------------------------------------------------------------------------------------/
    '/Step7.    FTP 서버의 해당 폴더에 자료가 있으면 최대 확장자 값 찾기/자료가 없으면 0
    '/----------------------------------------------------------------------------------------------------/
    Call MMSFTP.FtpScanDirectory(strIMGFILEPATH, strFileName & "*.*")
    
    nImageSeq = 0
    If UBound(FtpScanFileName) > 0 Then
        For intX = 1 To UBound(FtpScanFileName)
            strTemp1 = Split(FtpScanFileName(intX), "@")
            If UBound(strTemp1) = 3 Then
                If Val(Left(strTemp1(3), InStr(strTemp1(3), ".") - 1)) > nImageSeq Then
                    nImageSeq = Val(Left(strTemp1(3), InStr(strTemp1(3), ".") - 1))
                End If
            End If
        Next intX
    End If
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step8.    선택된 미전송 자료 Rename 하면서 전송
    '/----------------------------------------------------------------------------------------------------/
    For intX = 1 To spr미전송.MaxRows
        If GET_CELL(spr미전송, 1, intX) = "1" Then
            nImageSeq = nImageSeq + 1
            rtn = MMSFTP.FTPUploadFile(gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, intX), strIMGFILEPATH & strFileName & Format(nImageSeq, "000") & Mid(GET_CELL(spr미전송, 2, intX), InStr(GET_CELL(spr미전송, 2, intX), ".")))
        End If
    Next intX
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step9.    FTP 해제
    '/----------------------------------------------------------------------------------------------------/
    Call MMSFTP.CloseConnection
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step10.   모두 전송 성공 시 HIS에 검사완료 Falg Update 실행.
    '/----------------------------------------------------------------------------------------------------/
    If OpenDB(gstrREG_DB_CONSTR) = False Then End
    
    ADC.BeginTrans
    
    '/HIS 최종결과 Update
    '''If FUNC_HIS_RST_UPDATE = False Then ADC.RollbackTrans: Call CloseDB: End
    '/원처방:   TMPRSCINFN
    gstrQuy = "UPDATE TMPRSCINFN SET "
    gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '560', "                      '/560.임시결과 이상
    gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "  '/최종수정자 ID
    gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSTIMESTAMP "                '/최종수정일자
    gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(fstr처방일자), "YYYYMMDD") & "' "  '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & Val(fstr처방SEQ) & " "                                     '/처방번호
    gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & fstr병록번호 & "' "                                   '/병록번호
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & fstr처방코드 & "' "                                   '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                     '/원처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                     '/처방History 번호
    If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

    '/실시처방: TMPRSCEXCN
    gstrQuy = "UPDATE TMPRSCEXCN SET "
    gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '560', "                      '/560.임시결과 이상
    gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "  '/최종수정자 ID
    gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSTIMESTAMP "                '/최종수정일자
    gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(fstr처방일자), "YYYYMMDD") & "' "  '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & Val(fstr처방SEQ) & " "                                     '/처방번호
    gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & fstr병록번호 & "' "                                   '/병록번호
    gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & fstr처방코드 & "' "                                   '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                     '/실시처방 살아있는 처방
    If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    

    '/----------------------------------------------------------------------------------------------------/
    '/Step11.   모두 전송 성공 시 MM_EMR_RES(Image 결과 정보)에 Insert
    '/----------------------------------------------------------------------------------------------------/
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & fstr병록번호 & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(fstr처방일자, "-", "") & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(fstr처방SEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                
    If Not ADR Is Nothing Then
        ADR.Close: Set ADR = Nothing
        
        '/Server DB에 결과가 입력이 되어 있으면 검사일자만 Update 함.
        gstrQuy = "UPDATE MM_EMR_RES SET "
        gstrQuy = gstrQuy & vbCrLf & "       EXAMDATE  = TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD') "
        gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & fstr병록번호 & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(fstr처방일자, "-", "") & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(fstr처방SEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    Else
        '/장비코드별 처방코드 가져오기
        gstrQuy = "INSERT INTO MM_EMR_RES "
        gstrQuy = gstrQuy & vbCrLf & " (PATNO,      ORDDATE,    ORDSEQ,     EXAMDATE,       DEPTCODE, "
        gstrQuy = gstrQuy & vbCrLf & "  PARTCODE,   EQUIPCODE,  EXAMCODE,   WORDNO,         ROOMNO, "
        gstrQuy = gstrQuy & vbCrLf & "  IOFLAG,     EXECID,     DRID,       IMGFILENAME,    IMGFILEPATH, "
        gstrQuy = gstrQuy & vbCrLf & "  RECEDATE,   RECESEQ,    EQUIPSEQ) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & fstr병록번호 & "', "                    '/PATNO(병록번호)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Replace(fstr처방일자, "-", "") & "', "  '/ORDDATE(처방일자)
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(fstr처방SEQ) & ", "                 '/ORDSEQ(처방SEQ(건강검진일 경우 접수번호))
        gstrQuy = gstrQuy & vbCrLf & "  TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD'), "       '/EXAMDATE(결과입력일자)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstr진료과 & "', "                      '/DEPTCODE(진료과코드)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/PARTCODE(진료실코드)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPCODE & "', "          '/EQUIPCODE(장비코드)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstr처방코드 & "', "                    '/EXAMCODE(검사코드)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/WORDNO(병동코드)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/ROOMNO(병실코드)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstr입외구분 & "', "                    '/IOFLAG(입원/외래/종건 구분)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypUSER.USERID & "', "                '/EXECID(직원번호)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/DRID(처방의번호)
        gstrQuy = gstrQuy & vbCrLf & "  '" & strFileName & "', "                    '/IMGFILENAME(결과이미지파일명)
        gstrQuy = gstrQuy & vbCrLf & "  '" & strIMGFILEPATH & "', "                 '/IMGFILEPATH(결과이미지파일경로)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECEDATE(접수일자)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECESEQ(접수SEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPSEQ & "') "           '/EQUIPSEQ(장비SEQ)
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    End If
    
    ADC.CommitTrans
    
    Call CloseDB

    FUNC_MM_SAVE = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    MsgBox "전송 시 오류가 발생하였습니다." & vbCrLf & _
           "전산실 혹은 공급업체에 연락주시기 바랍니다.", vbCritical, "전송오류"

End Function

Public Function FUNC_MM_VIEW(ArgSection As Integer) As Boolean
    Dim str처방코드     As String
    
    FUNC_MM_VIEW = False
    
    Select Case ArgSection
        Case 1: GoSub VIEW1_RTN '/수진자내역(미전송)
        Case 2: GoSub VIEW2_RTN '/수진자내역(전송)
        Case 3: GoSub VIEW3_RTN '/미전송 자료불러오기
        Case 4: GoSub VIEW4_RTN '/전송완료 자료불러오기
    End Select
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

VIEW1_RTN: '/수진자내역(미전송)
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/장비코드별 처방코드 가져오기
        gstrQuy = "SELECT ORDCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            Do Until ADR.EOF
                str처방코드 = str처방코드 & ",'" & Trim(ADR!ORDCD & "") & "'"
                    
                ADR.MoveNext
            Loop
            ADR.Close: Set ADR = Nothing
                        
            str처방코드 = Mid(str처방코드, 2)
        End If
        
        If FUNC_HIS_ORDER1_VIEW(str처방코드) = True Then '/병원별 미전송 처방조회Query
            If ReadSQL(gstrQuy, ADR) = True Then
                If Not ADR Is Nothing Then
                    Do Until ADR.EOF
                        With spr수진자내역
                            .MaxRows = .MaxRows + 1: .Row = .MaxRows
                            
                            .Col = 2:  .Text = Trim(ADR!CHRTNO & "")                            '/병록번호
                            .Col = 3:  .Text = Trim(ADR!IO_SECTION & "")                        '/환자구분
                            .Col = 4:  .Text = Trim(ADR!DETPCD & "")                            '/진료과
                            .Col = 5:  .Text = Trim(ADR!PATNM & "")                             '/수진자명
                            .Col = 6:  .Text = Trim(ADR!SEX & "") & "/" & Trim(ADR!AGE & "")    '/Seq/Age
                            .Col = 7:  .Text = Format(Trim(ADR!ORDDATE & ""), "@@@@-@@-@@")     '/처방일자
                            .Col = 8:  .Text = Trim(ADR!ORDSEQ & "")                            '/처방SEQ
                            .Col = 9:  .Text = ""                                               '/결과일자(전송자료만)
                            .Col = 10: .Text = Trim(ADR!ORDNM & "")                             '/처방명
                            .Col = 11: .Text = Trim(ADR!ORDCD & "")                             '/처방코드
                            .Col = 12: .Text = Trim(ADR!CNDT_PRSC_STAT_CD & "")                 '/실시처방 처방진행상태Flag
                        End With
                        
                        ADR.MoveNext
                    Loop
                    ADR.Close: Set ADR = Nothing
                    
                End If
            End If
        End If
        
        Call CloseDB
    End If
Return

'/----------------------------------------------------------------------------------------------------/

VIEW2_RTN: '/수진자내역(전송)
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/장비코드별 처방코드 가져오기
        gstrQuy = "SELECT ORDCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            Do Until ADR.EOF
                str처방코드 = str처방코드 & ",'" & Trim(ADR!ORDCD & "") & "'"
                    
                ADR.MoveNext
            Loop
            ADR.Close: Set ADR = Nothing
                        
            str처방코드 = Mid(str처방코드, 2)
        End If
        
        If FUNC_HIS_ORDER2_VIEW(str처방코드) = True Then '/병원별 기전송 처방조회Query
            If ReadSQL(gstrQuy, ADR) = True Then
                If Not ADR Is Nothing Then
                    Do Until ADR.EOF
                        With spr수진자내역
                            .MaxRows = .MaxRows + 1: .Row = .MaxRows
                            
                            .Col = 2:  .Text = Trim(ADR!CHRTNO & "")                            '/병록번호
                            .Col = 3:  .Text = Trim(ADR!IO_SECTION & "")                        '/환자구분
                            .Col = 4:  .Text = Trim(ADR!DETPCD & "")                            '/진료과
                            .Col = 5:  .Text = Trim(ADR!PATNM & "")                             '/수진자명
                            .Col = 6:  .Text = Trim(ADR!SEX & "") & "/" & Trim(ADR!AGE & "")    '/Seq/Age
                            .Col = 7:  .Text = Format(Trim(ADR!ORDDATE & ""), "@@@@-@@-@@")     '/처방일자
                            .Col = 8:  .Text = Trim(ADR!ORDSEQ & "")                            '/처방SEQ
                            .Col = 9:  .Text = Format(Trim(ADR!EXAMDATE & ""), "@@@@-@@-@@")    '/결과일자(전송자료만)
                            .Col = 10: .Text = Trim(ADR!ORDNM & "")                             '/처방명
                            .Col = 11: .Text = Trim(ADR!ORDCD & "")                             '/처방코드
                            .Col = 12: .Text = Trim(ADR!CNDT_PRSC_STAT_CD & "")                 '/실시처방 처방진행상태Flag
                        End With
                    
                        ADR.MoveNext
                    Loop
                    ADR.Close: Set ADR = Nothing
                End If
            End If
        End If
        
        Call CloseDB
    End If
Return

'/----------------------------------------------------------------------------------------------------/

VIEW3_RTN:
    Dim MyPath, MyName
    Dim sFileCnt As Integer
    
    If spr미전송.MaxRows > 0 Then spr미전송.MaxRows = 0
    
    MyPath = gtypEQ_INFO.EQIMGFILEPATH & "\"   ' 경로를 설정합니다.
    MyName = Dir(MyPath, vbDirectory)   ' 첫번째 항목을 검색합니다.
    
'''    '/TIF or TIFF 파일 있으면 JPG 변환 후 삭제
'''    Do While MyName <> ""   ' 루프(loop)를 시작합니다.
'''        ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
'''        If MyName <> "." And MyName <> ".." Then
'''            If InStr(UCase(MyName), ".TIF") > 0 Or InStr(UCase(MyName), ".TIFF") > 0 Then
'''
'''                staCondition.Panels.Item(2).Text = "TIF or TIFF 파일을 JPG 로 변환 중입니다..."
'''
'''                sFileCnt = FUNC_TifToJpg(gtypEQ_INFO.EQIMGFILEPATH, CStr(MyName))
'''
'''                If sFileCnt = 0 Then
'''                    MsgBox "TIF 혹은 TIFF 파일을 JPG로 변환 중에 오류가 발생하였습니다." & vbCrLf & _
'''                           "전산실 혹은 공급업체에 연락주시기 바랍니다.", vbCritical, "자료불러오기 오류"
'''                End If
'''            End If
'''        End If
'''        MyName = Dir   ' 다음 항목을 읽어들입니다.
'''    Loop
'''    staCondition.Panels.Item(2).Text = ""
    
    MyName = Dir(MyPath, vbDirectory)   ' 첫번째 항목을 검색합니다.
    Do While MyName <> ""   ' 루프(loop)를 시작합니다.
       ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
        If MyName <> "." And MyName <> ".." Then
            If InStr(UCase(MyName), ".JPG") > 0 Or InStr(UCase(MyName), ".JPEG") > 0 Or _
               InStr(UCase(MyName), ".TIF") > 0 Or InStr(UCase(MyName), ".TIFF") > 0 Or _
               InStr(UCase(MyName), ".PDF") > 0 Then
                spr미전송.MaxRows = spr미전송.MaxRows + 1
                
                Call SET_CELL(spr미전송, 2, spr미전송.MaxRows, MyName) '/화일명(이름변경가능)
                Call SET_CELL(spr미전송, 3, spr미전송.MaxRows, FileDateTime(MyPath & MyName)) '/
            End If
        End If
        MyName = Dir   ' 다음 항목을 읽어들입니다.
    Loop
Return

'/----------------------------------------------------------------------------------------------------/

VIEW4_RTN:
    '/Step1.전송완료 자료 폴더 생성 및 폴더 Clear
    '/Step2.전송완료 DB 자료 조회
    '/Step3.FTP 서버자료 가져오기
    Dim strIMGFILEPATH  As String
    
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/----------------------------------------------------------------------------------------------------/
        '/Step1.전송완료 자료 폴더 생성 및 폴더 Clear
        '/----------------------------------------------------------------------------------------------------/
        If Dir(gtypEQ_INFO.FTPIMGFILEPATH, vbDirectory) = "" Then MkDir gtypEQ_INFO.FTPIMGFILEPATH
    
        If Len(Dir(gtypEQ_INFO.FTPIMGFILEPATH & "\*.*")) > 0 Then
            Kill gtypEQ_INFO.FTPIMGFILEPATH & "\*.*"
        End If
        
        '/----------------------------------------------------------------------------------------------------/
        '/Step2.전송완료 DB 자료 조회
        '/----------------------------------------------------------------------------------------------------/
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
        gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl병록번호 & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbl처방일자, "-", "") & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbl처방SEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            strIMGFILEPATH = Trim(ADR!IMGFILEPATH & "")                              '/FTP경로
            
            ADR.Close: Set ADR = Nothing
        End If
            
        Call CloseDB
        
        '/----------------------------------------------------------------------------------------------------/
        '/Step3.FTP 서버자료 가져오기
        '/----------------------------------------------------------------------------------------------------/
        '/FTP 접속 시도
        Dim success As Long
        success = sftp.IsConnected
        If (success <> 1) Then
            If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
                MsgBox "Image File Server에 접근할 수 없습니다." & vbCrLf & "전산실에 문의바랍니다.", vbCritical, "FTP 접속 실패"
                Exit Function
            End If
        End If
        
        '/기준 폴더 변경
        If MMSFTP.SetFTPDirectory(strIMGFILEPATH) = False Then
            '''MsgBox "해당 자료가 없습니다.", vbInformation, "확인"
            Exit Function
        End If
        
        '/오더SEQ 관련 파일 찾기
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
        
        With spr전송완료
            If UBound(FtpScanFileName_IMG) > 0 Then
                For intX = 1 To UBound(FtpScanFileName_IMG)
                    .MaxRows = .MaxRows + 1: .Row = .MaxRows
                    
                    .Col = 2: .Text = FtpScanFileName_IMG(intX)
                    '''.Col = 3: .Text = FtpScanFileDate(intX) '/날짜 관계는 필요에 따라 나타낸다(결과일자가 있는 관계로 굳이 보여줄 필요없다)
                    
                    If .MaxTextRowHeight(.Row) > 13.3 Then .RowHeight(.Row) = .MaxTextRowHeight(.Row)
                    
                    If MMSFTP.FTPDownloadFile(gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, intX), strIMGFILEPATH & GET_CELL(spr전송완료, 2, intX)) = True Then FUNC_MM_VIEW = True
                    
                    '/Call sftp.DownloadFileByName(strIMGFILEPATH & GET_CELL(spr전송완료, 2, intX), gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, intX))
                    
                    
                    FUNC_MM_VIEW = True
                Next intX
            End If
        End With
        
        '/FTP 접속 해재
        Call MMSFTP.CloseConnection
    End If
Return
End Function

Public Function FUNC_TifToJpg(ArgFilePath As String, ArgFileName As String) As Integer
'''    'tif 파일을 jpg 파올로 변환하는 함수
'''    Dim strFileExt        As String
'''    Dim strFilePath       As String
'''    Dim strFileName       As String
'''    Dim strFileNewName    As String
'''    Dim xPoint          As Integer
'''    Dim iImgCnt         As Integer
'''    Dim jCreateImgCnt   As Integer
'''
'''    FUNC_TifToJpg = 0
'''
'''On Error GoTo ERR_RTN
'''
'''    strFilePath = ArgFilePath '파일경로
'''    strFileName = ArgFileName '파일명
'''    xPoint = InStr(1, strFileName, ".")
'''    strFileExt = Trim(Mid(strFileName, xPoint + 1)) '확장자명
'''    strFileNewName = Trim(Mid(strFileName, 1, xPoint - 1)) '확장자 제외한 파일명
'''
'''    imvTif.LoadMultiPage strFilePath & "\" & strFileName, 0 '파일로드
'''    iImgCnt = imvTif.GetTotalPage '로드된 Tiff 파일의 Page 수
'''
'''    If LCase(strFileExt) = "tiff" Or LCase(strFileExt) = "tif" Then
'''        For jCreateImgCnt = 1 To iImgCnt '페이지 수만큼 jpg 파일 생성
'''            imvTif.ExportTIF strFilePath & "\" & strFileName, strFilePath & "\" & strFileNewName & "-" & Format(jCreateImgCnt, "00"), "JPG", jCreateImgCnt, 1
'''        Next jCreateImgCnt
'''        imvTif.Filename = ""
'''        '' jpg파일 생성후 tiff 파일 삭제
'''        Kill strFilePath & "\" & strFileName
'''    End If
'''    '생성된 jpg 파일 갯수 반환
'''    FUNC_TifToJpg = jCreateImgCnt
'''
'''ERR_RTN:
'''
End Function

Private Sub cboZoomValue_Click()
    Select Case cboZoomValue
        Case "25%": imvResult.View = 1
        Case "33%": imvResult.View = 2
        Case "50%": imvResult.View = 3
        Case "75%": imvResult.View = 4
        Case "100%": imvResult.View = 5
        Case "150%": imvResult.View = 6
        Case "200%": imvResult.View = 7
        Case Else: imvResult.View = 9
    End Select
    
    '/서북병원 상위버전 이미지 OCX에는 ZOOM 모듈이 없다.
'''    imvResult.Zoom Val(Replace(cboZoomValue, "%", "")), Val(Replace(cboZoomValue, "%", ""))
End Sub

Private Sub cboZoomValue_KeyDown(KeyCode As Integer, Shift As Integer)
    '/서북병원 상위버전 이미지 OCX에는 ZOOM 모듈이 없다.
'''    If KeyCode = vbKeyReturn Then
'''        strTemp = Replace(cboZoomValue, "%", "")
'''
'''        If IsNumeric(strTemp) = False Then
'''            MsgBox "1에서 200까지의 숫자를 (재)입력하십시오!", vbCritical, "비율조정불가"
'''            cboZoomValue.SetFocus
'''            Exit Sub
'''        End If
'''
'''        If Not (Val(strTemp) >= 1 And Val(strTemp) <= 200) Then
'''            MsgBox "1에서 200까지의 숫자를 (재)입력하십시오!", vbCritical, "비율조정불가"
'''            cboZoomValue.SetFocus
'''            Exit Sub
'''        End If
'''
'''        imvResult.Zoom Val(strTemp), Val(strTemp)
'''    End If
End Sub

Private Sub cmd미전송폴더변경_Click()
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    
    Message = "수정할 미전송폴더경로를 입력하십시오"   ' 프롬프트 설정.
    Title = "미전송 폴더 변경"   ' 제목 설정.
    ' 메시지 화면 표시, 제목, 기본값.
    MyValue = InputBox(Message, Title, gtypEQ_INFO.EQIMGFILEPATH)
    
    If Trim(MyValue) <> "" Then
        '/변경폴더 유무 확인
        If Dir(MyValue, vbDirectory) = "" Then '/변경폴더가 없으면...
            MsgBox "변경된 미전송 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "변경을 하려면 [미전송폴더변경]을 (재)실행하십시오.", vbCritical, "변경실패"
            Exit Sub
        Else
            '/2.변수설정
            gtypEQ_INFO.EQIMGFILEPATH = MyValue
        End If
        
        '/3.DB 저장
        If OpenDB(gstrREG_DB_CONSTR) = False Then End
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            ADR.Close: Set ADR = Nothing
            
            ADC.BeginTrans
    
            gstrQuy = "UPDATE MM_EMR_CONF SET "
            gstrQuy = gstrQuy & vbCrLf & "       EQIMGFILEPATH  = '" & gtypEQ_INFO.EQIMGFILEPATH & "' " '/Local 이미지 경로
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ       =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            ADC.CommitTrans
        End If
        
        Call CloseDB
    End If
End Sub

Private Sub cmd전송완료폴더변경_Click()
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    
    Message = "수정할 전송완료폴더경로를 입력하십시오"   ' 프롬프트 설정.
    Title = "전송완료 폴더 변경"   ' 제목 설정.
    ' 메시지 화면 표시, 제목, 기본값.
    MyValue = InputBox(Message, Title, gtypEQ_INFO.FTPIMGFILEPATH)
    
    If Trim(MyValue) <> "" Then
        '/변경폴더 유무 확인
        If Dir(MyValue, vbDirectory) = "" Then '/변경폴더가 없으면...
            MsgBox "변경된 전송완료 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "변경을 하려면 [전송완료폴더변경]을 (재)실행하십시오.", vbCritical, "변경실패"
            Exit Sub
        Else
            '/2.변수설정
            gtypEQ_INFO.FTPIMGFILEPATH = MyValue
        End If
    
        '/3.DB 저장
        If OpenDB(gstrREG_DB_CONSTR) = False Then End
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            ADR.Close: Set ADR = Nothing
            
            ADC.BeginTrans
    
            gstrQuy = "UPDATE MM_EMR_CONF SET "
            gstrQuy = gstrQuy & vbCrLf & "       FTPIMGFILEPATH = '" & gtypEQ_INFO.FTPIMGFILEPATH & "' " '/Local 이미지 경로
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ       =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            ADC.CommitTrans
        End If
        
        Call CloseDB
    End If
End Sub

Private Sub cmdDeleteFTP_Click()
    If Trim(gtypEQ_INFO.FTPIMGFILEPATH) = "" Then MsgBox "전송완료 파일 저장 폴더 정보가 없습니다.", vbCritical, "조회불가": Exit Sub
    
    strTemp = "N"
    For intX = 1 To spr전송완료.MaxRows
        If GET_CELL(spr전송완료, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "삭제할 파일을 선택하십시오!", vbCritical, "삭제불가": Exit Sub
    
    If GET_CELL(spr수진자내역, 12, spr수진자내역.ActiveRow) >= "610" Then
        MsgBox "해당 처방은 실시 완료된 처방입니다. 삭제할 수 없습니다.!" & vbCrLf & vbCrLf & _
               "실시 해제 후 삭제 바랍니다.", vbCritical, "삭제불가": Exit Sub
    End If
    
    If MsgBox("선택한 파일은 결과 처리가 완료된 상태이며, " & vbCrLf & _
              "삭제 시 복구가 불가능하오니 주의하시기 바랍니다." & vbCrLf & vbCrLf & _
              "계속해서 삭제 처리를 진행하겠습니까?선택한 전송완료 파일을 삭제하겠습니까?", vbQuestion + vbYesNo, "삭제질의") = vbNo Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    If FUNC_MM_DELETE("2") = True Then Call cmdView_Click 'Call cmdViewFTP_Click
End Sub

Private Sub cmdDeleteLocal_Click()
    If Trim(gtypEQ_INFO.EQIMGFILEPATH) = "" Then MsgBox "미전송 파일 저장 폴더 정보가 없습니다.", vbCritical, "삭제불가": Exit Sub
    
    strTemp = "N"
    For intX = 1 To spr미전송.MaxRows
        If GET_CELL(spr미전송, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "삭제할 파일을 선택하십시오!", vbCritical, "삭제불가": Exit Sub
    
    If MsgBox("선택한 미전송 파일을 삭제하겠습니까?", vbQuestion + vbOKCancel, "삭제질의") = vbCancel Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    If FUNC_MM_DELETE("1") = True Then Call cmdViewLocal_Click
End Sub

Private Sub cmdMultiFirst_Click()
    txtMultiPno = 1
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiJump_Click()
    imvResult.LoadMultiPage imvResult.Filename, Val(txtMultiPno)
End Sub

Private Sub cmdMultiLast_Click()
    txtMultiPno = imvResult.GetTotalPage
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiNext_Click()
    If txtMultiPno < imvResult.GetTotalPage Then
        txtMultiPno = txtMultiPno + 1
    Else
        txtMultiPno = imvResult.GetTotalPage
    End If
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiPrev_Click()
    If txtMultiPno > 1 Then
        txtMultiPno = txtMultiPno - 1
    Else
        txtMultiPno = 1
    End If
    Call cmdMultiJump_Click
End Sub

Private Sub cmdRotate_Click()
    imvResult.Rotate90
    imvResult.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim intRow수진자내역    As Integer
    
    '/----------------------------------------------------------------------------------------------------/
    '/## [Image List가 전송완료 Tab일 경우 전송버튼은 Enabled = False 가 된다.
    '/Step1.    수진자내역에 하나 이상의 환자가 선택되어 있는지 확인
    '/Step2.    수진자내역에 하나 이상의 환자가 선택되어 있을때 해당 환자의 병록번호가 다른지 확인
    '/Step3.    [Image List]의 미전송 Tab에 하나 이상의 자료가 선택되었는지 확인
    '/Step4.    FTP 연결
    '/Step5.    FTP 서버의 해당 폴더 생성/ 있으면 SKIP
    '/Step6.    전송할 확장자 제외한 파일명 정의
    '/Step7.    FTP 서버의 해당 폴더에 자료가 있으면 최대 확장자 값 찾기/자료가 없으면 0
    '/Step8.    선택된 미전송 자료 Rename 하면서 전송
    '/Step9.    FTP 해제
    '/Step10.   모두 전송 성공 시 HIS에 검사완료 Falg Update 실행.
    '/Step11.   모두 전송 성공 시 MM_EMR_RES(Image 결과 정보)에 Insert
    '/Step12.   모두 전송 성공 시 실제 자료를 Local폴더에서 삭제 후 수진자내역 조회 및 미전송 자료불러오기 실행
    '/----------------------------------------------------------------------------------------------------/
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step1.    수진자내역에 하나 이상의 환자가 선택되어 있는지 확인
    '/----------------------------------------------------------------------------------------------------/
    strTemp = "N"
    For intX = 1 To spr수진자내역.MaxRows
        If GET_CELL(spr수진자내역, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "전송할 수진자를 선택하십시오!", vbCritical, "전송불가": Exit Sub

    '/----------------------------------------------------------------------------------------------------/
    '/Step2.    수진자내역에 하나 이상의 환자가 선택되어 있을때 해당 환자의 병록번호가 다른지 확인
    '/----------------------------------------------------------------------------------------------------/
    strTemp = ""
    For intX = 1 To spr수진자내역.MaxRows
        If GET_CELL(spr수진자내역, 1, intX) = "1" Then
            If strTemp <> GET_CELL(spr수진자내역, 2, intX) Then
                If strTemp = "" Then
                    strTemp = GET_CELL(spr수진자내역, 2, intX)
                Else
                    MsgBox "2개 이상 선택한, 전송할 수진자의 병록번호가 전혀 다릅니다.", vbCritical, "전송불가": Exit Sub
                End If
            End If
        End If
    Next intX

    '/----------------------------------------------------------------------------------------------------/
    '/Step3.    [Image List]의 미전송 Tab에 하나 이상의 자료가 선택되었는지 확인
    '/----------------------------------------------------------------------------------------------------/
    strTemp = "N"
    For intX = 1 To spr미전송.MaxRows
        If GET_CELL(spr미전송, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "전송할 Image를 선택하십시오!", vbCritical, "전송불가": Exit Sub

    On Error GoTo ERN_ERR
    
    Screen.MousePointer = 11
    
    For intRow수진자내역 = 1 To spr수진자내역.MaxRows
        If GET_CELL(spr수진자내역, 1, intRow수진자내역) = "1" Then
            If FUNC_MM_SAVE(intRow수진자내역) = False Then GoTo ERN_ERR
        End If
    Next intRow수진자내역
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step12.   모두 전송 성공 시 실제 자료를 Local폴더에서 삭제 후 수진자내역 조회 및 미전송 자료불러오기 실행
    '/----------------------------------------------------------------------------------------------------/
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")
    For intX = 1 To spr미전송.MaxRows
        If GET_CELL(spr미전송, 1, intX) = "1" Then
            Kill gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, intX)
        End If
    Next intX
    Call cmdView_Click
    Call cmdViewLocal_Click
    
    Screen.MousePointer = 0
    
    MsgBox "전송되었습니다.", vbInformation, "확인"
    
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ERN_ERR:
    Screen.MousePointer = 0
End Sub

Private Sub cmdView_Click()
    Call FUNC_MM_KEY_CLEAR("1") '/수진자내역 Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage전송여부(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/전송완료 Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
    
    Select Case True
        Case opt전송여부(0).Value
            Call FUNC_MM_VIEW("1") '/수진자내역(미전송)
            If spr수진자내역.MaxRows > 0 Then Call spr수진자내역_LeaveCell(0, 0, 1, 1, False)
            
        Case opt전송여부(1).Value
            Call FUNC_MM_VIEW("2") '/수진자내역(전송)
            If spr수진자내역.MaxRows > 0 Then Call spr수진자내역_LeaveCell(0, 0, 1, 1, False)
            
            If optImage전송여부(1).Value = True Then Call cmdViewFTP_Click
    End Select
End Sub

Private Sub cmdViewFTP_Click()
    Call FUNC_MM_KEY_CLEAR("5") '/전송완료 Spread Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    If Trim(gtypEQ_INFO.FTPIMGFILEPATH) = "" Then '/DB에서 가져온 FTP결과파일저장경로가 없을 때...
        If MsgBox("전송완료 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "기본 폴더로 지정하겠습니까?", vbQuestion + vbOKCancel, "전송완료폴더질의") = vbCancel Then Exit Sub
        
        GoSub RTN_기본폴더생성및설정
    
    Else '/DB에서 가져온 FTP결과파일저장경로가 있을 때...
        '/기존폴더 유무 확인
        If Dir(gtypEQ_INFO.FTPIMGFILEPATH, vbDirectory) = "" Then '/기존폴더가 없으면...
            MsgBox "기존에 설정된 전송완료 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "변경을 하려면 [전송완료폴더변경]을 실행하십시오.", vbInformation, "확인"
            Exit Sub
        End If
    End If
    If spr수진자내역.MaxRows > 0 And Trim(lbl병록번호) = "" Then MsgBox "조회할 수진자를 선택하십시오!", vbCritical, "조회불가": Exit Sub
    
    Call FUNC_MM_VIEW("4") '/전송완료 자료불러오기
    If spr전송완료.MaxRows > 0 Then
        imvResult.Filename = ""
        imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, spr전송완료.ActiveRow)
        Call SUB_CHK_PDF_TIF(imvResult.Filename)
        imvResult.View = 9
    End If
Exit Sub

'/----------------------------------------------------------------------------------------------------/

RTN_기본폴더생성및설정:
    '/1.기본폴더 생성
    If SET_DEFAULT_FOLDER("FTP") = False Then
        MsgBox "기본폴더가 생성되지 않았습니다." & vbCrLf & vbCrLf & _
               "전산실 혹은 공급업체에 문의하시기 바랍니다.", vbCritical, "기본폴더생성실패"
               
        Exit Sub
    End If
    
    '/2.변수설정
    gtypEQ_INFO.FTPIMGFILEPATH = App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ & "\" & "EQUIP"
Return
End Sub

Private Sub cmdViewLocal_Click()
    Call FUNC_MM_KEY_CLEAR("4") '/미전송 Spread Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    If Trim(gtypEQ_INFO.EQIMGFILEPATH) = "" Then '/DB에서 가져온 장비결과파일저장경로가 없을 때...
        If MsgBox("미전송 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "기본 폴더로 지정하겠습니까?", vbQuestion + vbOKCancel, "미전송폴더질의") = vbCancel Then Exit Sub
        
        GoSub RTN_기본폴더생성및설정
        
    Else '/DB에서 가져온 장비결과파일저장경로가 있을 때...
        '/기존폴더 유무 확인
        If Dir(gtypEQ_INFO.EQIMGFILEPATH, vbDirectory) = "" Then '/기존폴더가 없으면...
            MsgBox "기존에 설정된 미전송 폴더가 존재하지 않습니다." & vbCrLf & vbCrLf & _
                   "변경을 하려면 [미전송폴더변경]을 실행하십시오.", vbInformation, "확인"
            Exit Sub
        End If
    End If
    
    Call FUNC_MM_VIEW("3") '/미전송 자료불러오기
    If spr미전송.MaxRows > 0 Then
        imvResult.Filename = ""
        imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, spr미전송.ActiveRow)
        Call SUB_CHK_PDF_TIF(imvResult.Filename)
        imvResult.View = 9
    End If
Exit Sub

'/----------------------------------------------------------------------------------------------------/

RTN_기본폴더생성및설정:
    '/1.기본폴더 생성
    If SET_DEFAULT_FOLDER("EQUIP") = False Then
        MsgBox "기본폴더가 생성되지 않았습니다." & vbCrLf & vbCrLf & _
               "전산실 혹은 공급업체에 문의하시기 바랍니다.", vbCritical, "기본폴더생성실패"
               
        Exit Sub
    End If
    
    '/2.변수설정
    gtypEQ_INFO.EQIMGFILEPATH = App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ & "\" & "EQUIP"
Return
End Sub

Private Sub cmdzoomin_Click()
'''    Dim strEXAMCODE As String
'''    Dim strUpEXAMCODE As String
'''    Dim nInc        As Integer
'''
'''    If OpenDB(gstrREG_DB_CONSTR) = True Then
'''        ADC.BeginTrans
'''
'''        gstrQuy = "SELECT EXAMCODE, ROWID "
'''        gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMMASTER_1 "
'''        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXAMCODE "
'''        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
'''
'''        If Not ADR Is Nothing Then
'''            Do Until ADR.EOF
'''                If strEXAMCODE = Trim(ADR!EXAMCODE & "") Then
'''                    nInc = nInc + 1
'''
'''                    strUpEXAMCODE = Trim(ADR!EXAMCODE & "") & "_" & CStr(nInc)
'''                Else
'''                    nInc = 1
'''
'''                    strUpEXAMCODE = Trim(ADR!EXAMCODE & "") & "_" & CStr(nInc)
'''
'''                    strEXAMCODE = Trim(ADR!EXAMCODE & "")
'''                End If
'''
'''                gstrQuy = "UPDATE EXAMMASTER_1 SET "
'''                gstrQuy = gstrQuy & vbCrLf & "       EXAMCODE1 = '" & strUpEXAMCODE & "' "
'''                gstrQuy = gstrQuy & vbCrLf & " WHERE ROWID  = '" & Trim(ADR!ROWID & "") & "' "
'''                If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
'''
'''                ADR.MoveNext
'''            Loop
'''            ADR.Close: Set ADR = Nothing
'''        End If
'''
'''        ADC.CommitTrans
'''
'''        Call CloseDB
'''    End If
    imvResult.ZoomIn
End Sub

Private Sub cmdzoomout_Click()
    imvResult.ZoomOut
End Sub

Private Sub cmdZmHeight_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 11
    imvResult.SetFocus
End Sub

Private Sub cmdFit_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 9
    imvResult.SetFocus
End Sub

Private Sub cmdZmWidth_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 10
    imvResult.SetFocus
End Sub

Private Sub cmdCenter_Click()
    cmdzoomin.Enabled = False
    cmdzoomout.Enabled = False
    cboZoomValue.Enabled = False
    cmdRotate.Enabled = False
    
    imvResult.View = 12
    imvResult.SetFocus
End Sub

Private Sub cmd100_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 5
    imvResult.SetFocus
End Sub

Private Sub dtp접수일자_Change()
    Call FUNC_MM_KEY_CLEAR("1") '/수진자내역 Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage전송여부(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/전송완료 Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
End Sub

Private Sub dtp접수일자_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown, Txt
   
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
   
    If KeyCode = vbKeyM Then   ' 키의 조합 상태를 출력합니다.
        If mnuSetting.Visible = True Then
            mnuSetting.Visible = False
        Else
            mnuSetting.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Call FUNC_MM_INITIAL
    
    DoEvents
    DoEvents
    DoEvents
    
    Call cmdView_Click
    
    DoEvents
    DoEvents
    DoEvents
    
    Call cmdViewLocal_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : 높이가 늘어나는 개체 3개, 디자인상 해당 개체 위에 늘어난 개체가 2개
    For intX = 0 To UBound(CW)
        Select Case CW(intX).Nm
            Case fraLimageList.Name:        fraLimageList.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case picImageList미전송.Name:   picImageList미전송.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case picImageList전송완료.Name: picImageList전송완료.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case spr미전송.Name:            spr미전송.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case spr전송완료.Name:          spr전송완료.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case cmdViewLocal.Name:         cmdViewLocal.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmd미전송폴더변경.Name:    cmd미전송폴더변경.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdDeleteLocal.Name:       cmdDeleteLocal.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdViewFTP.Name:           cmdViewFTP.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmd전송완료폴더변경.Name:  cmd전송완료폴더변경.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdDeleteFTP.Name:         cmdDeleteFTP.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdSave.Name:              cmdSave.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case shpPatientInfo.Name:       shpPatientInfo.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case prgPatient.Name:           prgPatient.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picImage.Name:             picImage.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height + (Me.Height - lngMeHeight)
            Case imvResult.Name:            imvResult.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height + (Me.Height - lngMeHeight)
            Case picTifPdf.Name:            picTifPdf.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picControl.Name:           picControl.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picJPG.Name:               picJPG.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
        End Select
    Next intX
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MMSFTP.CloseConnection

    Set MMSFTP = Nothing
    Call CloseDB
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")
'''    imvTif.Filename = ""

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents

    '/FTP결과파일저장경로 폴더의 모든 파일 삭제
    If Len(Dir(gtypEQ_INFO.FTPIMGFILEPATH & "\*.*")) > 0 Then
        Kill gtypEQ_INFO.FTPIMGFILEPATH & "\*.*"
    End If

    Set frmVPM_Main = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frm공용_Info.Show vbModal
End Sub

Private Sub mnuSettingSub_Click(Index As Integer)
    Select Case Index
        Case 0: frm공용_Set_DataBase.Show vbModal
        Case 1: frm공용_Set_Equipment_List.Show vbModal
        Case 2: frm공용_Set_Equip_Config.Show vbModal
    End Select
End Sub

Private Sub MSComm1_OnComm()
    '/EMR Interface 대상 장비 중 장비에서 나오는 신호가 SM일때
    ' 출력할 수 있는 형태(Spread 등)로 만들어 가상프린터로 출력 시킨다.
    Select Case gtypEQ_INFO.EQUIPCODE
        Case "00008" '/AL2000(안과장비)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1) 'SOH
                    gstrMSCOMM_Buff = ""
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4) 'EOT
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00008.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
        
        Case "00014" '/렌즈미터(안과장비)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case vbCr
                
                Case vbLf
                    If Mid(gstrMSCOMM_Buff, 1, 5) = "LM2RK" Then
                        txtSerialData = gstrMSCOMM_Buff
                        Call frmVPM_SM00014.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    End If
                    gstrMSCOMM_Buff = ""
                    
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
        
        Case "00016" '/CT80(안과장비)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1)
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4)
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00016.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
    
        Case "00025" '/KR7100(안과장비)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1)
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4)
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00025.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
    End Select
End Sub

Private Sub opt전송여부_Click(Index As Integer)
    Select Case Index
        Case 0:
            opt전송여부(0).ForeColor = RGB(0, 0, 255)
            opt전송여부(0).FontBold = True
            opt전송여부(1).ForeColor = RGB(0, 0, 0)
            opt전송여부(1).FontBold = False
        Case 1:
            opt전송여부(0).ForeColor = RGB(0, 0, 0)
            opt전송여부(0).FontBold = False
            opt전송여부(1).ForeColor = RGB(0, 0, 255)
            opt전송여부(1).FontBold = True
    End Select
    
    Call FUNC_MM_KEY_CLEAR("1") '/수진자내역 Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage전송여부(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/전송완료 Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
End Sub

Private Sub opt전송여부_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub optImage전송여부_Click(Index As Integer)
    Select Case Index
        Case 0: '/미전송 Tab
            optImage전송여부(0).ForeColor = RGB(0, 0, 255)
            optImage전송여부(0).FontBold = True
            optImage전송여부(1).ForeColor = RGB(0, 0, 0)
            optImage전송여부(1).FontBold = False
            
            picImageList미전송.Visible = True
            picImageList전송완료.Visible = False
            
            cmdSave.Enabled = True
            
            Call cmdViewLocal_Click
            
            Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
            
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, spr미전송.ActiveRow)
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
            
        Case 1: '/전송완료 Tab
            optImage전송여부(0).ForeColor = RGB(0, 0, 0)
            optImage전송여부(0).FontBold = False
            optImage전송여부(1).ForeColor = RGB(0, 0, 255)
            optImage전송여부(1).FontBold = True
            
            picImageList미전송.Visible = False
            picImageList전송완료.Visible = True
            
            cmdSave.Enabled = False
            
            Call cmdViewFTP_Click
            
            Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
            
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, spr전송완료.ActiveRow)
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
    End Select

End Sub

Private Sub spr미전송_Click(ByVal Col As Long, ByVal Row As Long)
    With spr미전송
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr미전송, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr미전송.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr미전송, 1, intX, "0")
                Else
                    Call SET_CELL(spr미전송, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr미전송.Tag, 2)) = Col Then
                If Left(spr미전송.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr미전송.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr미전송.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr미전송.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr미전송_LeaveCell(0, 0, 1, spr미전송.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr미전송_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr미전송, 1, Row) = "1" Then
        Call SET_CELL(spr미전송, 1, Row, "0")
    Else
        Call SET_CELL(spr미전송, 1, Row, "1")
    End If
End Sub

Private Sub spr미전송_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr미전송, 1, spr미전송.ActiveRow) = "1" Then
            Call SET_CELL(spr미전송, 1, spr미전송.ActiveRow, "0")
        Else
            Call SET_CELL(spr미전송, 1, spr미전송.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr미전송_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub

    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    imvResult.Filename = ""
    imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, NewRow)
    Call SUB_CHK_PDF_TIF(imvResult.Filename)
    imvResult.View = 9
End Sub

Private Sub spr미전송_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    Dim strExtension    As String
    
    If Row < 1 Then Exit Sub
    
    Message = "수정할 파일명을 입력하십시오"   ' 프롬프트 설정.
    Title = "Image 파일명 변경"   ' 제목 설정.
    ' 메시지 화면 표시, 제목, 기본값.
    MyValue = InputBox(Message, Title, Left(GET_CELL(spr미전송, 2, Row), InStr(GET_CELL(spr미전송, 2, Row), ".") - 1))
    
    strExtension = Mid(GET_CELL(spr미전송, 2, Row), InStr(GET_CELL(spr미전송, 2, Row), ".") + 1)
    
    If Trim(MyValue) <> "" Then
        OldName = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr미전송, 2, Row)
        NewName = gtypEQ_INFO.EQIMGFILEPATH & "\" & MyValue & "." & strExtension  ' 파일 이름을 정의합니다.
        
        If Row = spr미전송.ActiveRow Then
            imvResult.Filename = ""
            Call SUB_CHK_PDF_TIF("")
        End If
        
        Name OldName As NewName   ' 파일 이름을 변경합니다.

        If Row = spr미전송.ActiveRow Then
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & MyValue & "." & strExtension
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
        End If
        
        Call SET_CELL(spr미전송, 2, Row, MyValue & "." & strExtension)
    End If
End Sub

Private Sub spr수진자내역_Click(ByVal Col As Long, ByVal Row As Long)
    With spr수진자내역
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr수진자내역, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr수진자내역.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr수진자내역, 1, intX, "0")
                Else
                    Call SET_CELL(spr수진자내역, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr수진자내역.Tag, 2)) = Col Then
                If Left(spr수진자내역.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr수진자내역.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr수진자내역.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr수진자내역.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr수진자내역_LeaveCell(0, 0, 1, spr수진자내역.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr수진자내역_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr수진자내역, 1, Row) = "1" Then
        Call SET_CELL(spr수진자내역, 1, Row, "0")
    Else
        Call SET_CELL(spr수진자내역, 1, Row, "1")
    End If
End Sub

Private Sub spr수진자내역_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr수진자내역, 1, spr수진자내역.ActiveRow) = "1" Then
            Call SET_CELL(spr수진자내역, 1, spr수진자내역.ActiveRow, "0")
        Else
            Call SET_CELL(spr수진자내역, 1, spr수진자내역.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr수진자내역_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear

    lbl병록번호 = GET_CELL(spr수진자내역, 2, NewRow)
    lbl수진자명 = GET_CELL(spr수진자내역, 5, NewRow)
    lbl연령성별 = GET_CELL(spr수진자내역, 6, NewRow)
    lbl처방명 = GET_CELL(spr수진자내역, 10, NewRow)
    lbl처방코드 = GET_CELL(spr수진자내역, 11, NewRow)

    lbl진료과 = GET_CELL(spr수진자내역, 4, NewRow)
    lbl입외구분 = GET_CELL(spr수진자내역, 3, NewRow)

    lbl처방일자 = GET_CELL(spr수진자내역, 7, NewRow)
    lbl처방SEQ = GET_CELL(spr수진자내역, 8, NewRow)
    lbl결과일자 = GET_CELL(spr수진자내역, 9, NewRow)
    
    Select Case GET_CELL(spr수진자내역, 12, NewRow) '/실시처방 처방진행상태Flag
        Case "440": lbl처방상태 = "접수": lbl처방상태.ForeColor = RGB(0, 0, 0)
        Case "560": lbl처방상태 = "임시결과": lbl처방상태.ForeColor = RGB(0, 255, 0)
        Case "610": lbl처방상태 = "실시완료": lbl처방상태.ForeColor = RGB(255, 0, 0)
        Case Else:  lbl처방상태 = GET_CELL(spr수진자내역, 12, NewRow)
    End Select
    
    If optImage전송여부(1).Value = True Then Call cmdViewFTP_Click
End Sub

Private Sub spr전송완료_Click(ByVal Col As Long, ByVal Row As Long)
    With spr전송완료
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr전송완료, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr전송완료.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr전송완료, 1, intX, "0")
                Else
                    Call SET_CELL(spr전송완료, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr전송완료.Tag, 2)) = Col Then
                If Left(spr전송완료.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr전송완료.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr전송완료.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr전송완료.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr전송완료_LeaveCell(0, 0, 1, spr전송완료.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr전송완료_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr전송완료, 1, Row) = "1" Then
        Call SET_CELL(spr전송완료, 1, Row, "0")
    Else
        Call SET_CELL(spr전송완료, 1, Row, "1")
    End If
End Sub

Private Sub spr전송완료_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr전송완료, 1, spr전송완료.ActiveRow) = "1" Then
            Call SET_CELL(spr전송완료, 1, spr전송완료.ActiveRow, "0")
        Else
            Call SET_CELL(spr전송완료, 1, spr전송완료.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr전송완료_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub

    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    imvResult.Filename = ""
    imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr전송완료, 2, NewRow)
    Call SUB_CHK_PDF_TIF(imvResult.Filename)
    imvResult.View = 9
End Sub

Public Sub SUB_GET_REG_CLIENT_INFO()
    Dim strEQCD             As String
    Dim strEQNM             As String
    Dim strEQSEQ            As String
    Dim strEQPOS            As String
    Dim strEQTYPE           As String
    Dim strRECEIVETYPE      As String
    Dim strEQUIPPORT        As String
    Dim strORDYN            As String
    Dim strQUERYTYPE        As String
    Dim strZIPYN            As String
    Dim strSERIALYN         As String
    Dim strSERIALPORT       As String
    Dim strSERIALBAUD       As String
    Dim strSERIALDATABIT    As String
    Dim strSERIALSTARTBIT   As String
    Dim strSERIALSTOPBIT    As String
    Dim strSERIALPARITY     As String
    Dim strSERIALRTS        As String
    Dim strSERIALDTR        As String
    Dim strEQIMGFILEPATH    As String
    Dim strFTPIMGFILEPATH   As String
    
    Dim strEQCD_Array
    Dim strEQNM_Array
    Dim strEQSEQ_Array
    Dim strEQPOS_Array
    Dim strEQTYPE_Array
    Dim strRECEIVETYPE_Array
    Dim strEQUIPPORT_Array
    Dim strORDYN_Array
    Dim strQUERYTYPE_Array
    Dim strZIPYN_Array
    Dim strSERIALYN_Array
    Dim strSERIALPORT_Array
    Dim strSERIALBAUD_Array
    Dim strSERIALDATABIT_Array
    Dim strSERIALSTARTBIT_Array
    Dim strSERIALSTOPBIT_Array
    Dim strSERIALPARITY_Array
    Dim strSERIALRTS_Array
    Dim strSERIALDTR_Array
    Dim strEQIMGFILEPATH_Array
    Dim strFTPIMGFILEPATH_Array
    
    '/대상 의료장비 정보(레지스터) 가져오기
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQNM)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)
    strEQPOS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQPOS)
    strEQTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQTYPE)
    strRECEIVETYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_RECEIVETYPE)
    strEQUIPPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQUIPPORT)
    strORDYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ORDYN)
    strQUERYTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_QUERYTYPE)
    strZIPYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPYN)
    strSERIALYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALYN)
    strSERIALPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPORT)
    strSERIALBAUD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALBAUD)
    strSERIALDATABIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDATABIT)
    strSERIALSTARTBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTARTBIT)
    strSERIALSTOPBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTOPBIT)
    strSERIALPARITY = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPARITY)
    strSERIALRTS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALRTS)
    strSERIALDTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDTR)
    strEQIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQIMGFILEPATH)
    strFTPIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_FTPIMGFILEPATH)
    
    strEQCD_Array = Split(strEQCD, ",")
    strEQNM_Array = Split(strEQNM, ",")
    strEQSEQ_Array = Split(strEQSEQ, ",")
    strEQPOS_Array = Split(strEQPOS, ",")
    strEQTYPE_Array = Split(strEQTYPE, ",")
    strRECEIVETYPE_Array = Split(strRECEIVETYPE, ",")
    strEQUIPPORT_Array = Split(strEQUIPPORT, ",")
    strORDYN_Array = Split(strORDYN, ",")
    strQUERYTYPE_Array = Split(strQUERYTYPE, ",")
    strZIPYN_Array = Split(strZIPYN, ",")
    strSERIALYN_Array = Split(strSERIALYN, ",")
    strSERIALPORT_Array = Split(strSERIALPORT, ",")
    strSERIALBAUD_Array = Split(strSERIALBAUD, ",")
    strSERIALDATABIT_Array = Split(strSERIALDATABIT, ",")
    strSERIALSTARTBIT_Array = Split(strSERIALSTARTBIT, ",")
    strSERIALSTOPBIT_Array = Split(strSERIALSTOPBIT, ",")
    strSERIALPARITY_Array = Split(strSERIALPARITY, ",")
    strSERIALRTS_Array = Split(strSERIALRTS, ",")
    strSERIALDTR_Array = Split(strSERIALDTR, ",")
    strEQIMGFILEPATH_Array = Split(strEQIMGFILEPATH, ",")
    strFTPIMGFILEPATH_Array = Split(strFTPIMGFILEPATH, ",")
    
    On Error Resume Next
    
    With sprEQ_INFO
        If .MaxRows > 0 Then .MaxRows = 0
        
        For intX = 0 To UBound(strEQCD_Array)
            .MaxRows = .MaxRows + 1: .Row = .MaxRows
        
            .Col = 1:   .Text = strEQCD_Array(intX)
            .Col = 2:   .Text = strEQNM_Array(intX)
            .Col = 3:   .Text = strEQSEQ_Array(intX)
            .Col = 4:   .Text = strEQPOS_Array(intX)
            .Col = 5:   .Text = strEQTYPE_Array(intX)
            .Col = 6:   .Text = strRECEIVETYPE_Array(intX)
            .Col = 7:   .Text = strEQUIPPORT_Array(intX)
            .Col = 8:   .Text = strORDYN_Array(intX)
            .Col = 9:   .Text = strQUERYTYPE_Array(intX)
            .Col = 10:  .Text = strZIPYN_Array(intX)
            .Col = 11:  .Text = strSERIALYN_Array(intX)
            .Col = 12:  .Text = strSERIALPORT_Array(intX)
            .Col = 13:  .Text = strSERIALBAUD_Array(intX)
            .Col = 14:  .Text = strSERIALDATABIT_Array(intX)
            .Col = 15:  .Text = strSERIALSTARTBIT_Array(intX)
            .Col = 16:  .Text = strSERIALSTOPBIT_Array(intX)
            .Col = 17:  .Text = strSERIALPARITY_Array(intX)
            .Col = 18:  .Text = strSERIALRTS_Array(intX)
            .Col = 19:  .Text = strSERIALDTR_Array(intX)
            .Col = 20:  .Text = strEQIMGFILEPATH_Array(intX)
            .Col = 21:  .Text = strFTPIMGFILEPATH_Array(intX)
        Next intX
    End With
    
    On Error GoTo 0
End Sub

Private Sub staCondition_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel = "COM" Then
        If txtSerialData.Visible = False Then
            txtSerialData.Visible = True
        Else
            txtSerialData.Visible = False
        End If
    End If
End Sub

Public Sub SUB_CHK_PDF_TIF(ArgPathFileName As String)
    If InStr(UCase(ArgPathFileName), ".TIF") > 0 Or InStr(UCase(ArgPathFileName), ".TIFF") > 0 Or InStr(UCase(ArgPathFileName), ".PDF") > 0 Then
        imvResult.LoadMultiPage ArgPathFileName, 1
        
        txtMultiPno = "1"
        lblMultiCnt = str(imvResult.GetTotalPage)
        
        picTifPdf.Visible = True
        picJPG.Visible = False
    Else
        lblMultiCnt = ""
        
        picTifPdf.Visible = False
        picJPG.Visible = True
    End If
End Sub

Public Function FUNC_HIS_RST_UPDATE() As Boolean
    
    FUNC_HIS_RST_UPDATE = False
    
    Select Case gstrHOS_CUSCD
        Case 1 '/1.인천의료원
            Select Case gtypEQ_INFO.QUERYTYPE
                Case "1" '/3내과
                    gstrQuy = "UPDATE SY_MEODPRSC SET "
                    gstrQuy = gstrQuy & vbCrLf & "       CDIS_YN            = 'Y' "
                    gstrQuy = gstrQuy & vbCrLf & " WHERE PID                = '" & lbl병록번호 & "' "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_DATE          = TO_DATE('" & dtp접수일자.Value & "','YYYY-MM-DD') "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_SQNO          = '" & lbl처방SEQ & "' "
                    If RunSQL(gstrQuy) = False Then Exit Function
        
                Case "3" '/종합검진
                    '/종합검진은 처방단위로 완료여부를 SETTING 할 곳이 없다.
        
                Case Else
                    gstrQuy = "UPDATE SY_MEODPRSC SET "
                    gstrQuy = gstrQuy & vbCrLf & "       CDIS_YN            = 'Y', "
                    gstrQuy = gstrQuy & vbCrLf & "       PRSC_PRGR_STAT_CD  ='C' "
                    gstrQuy = gstrQuy & vbCrLf & " WHERE PID                = '" & lbl병록번호 & "' "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_DATE          = TO_DATE('" & dtp접수일자.Value & "','YYYY-MM-DD') "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_SQNO          = '" & lbl처방SEQ & "' "
                    If RunSQL(gstrQuy) = False Then Exit Function
            End Select
            
        Case 2 '/2.서울시립서북병원
            '/원처방:   TMPRSCINFN
            gstrQuy = "UPDATE TMPRSCINFN SET "
            gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '560', "                      '/560.임시결과 이상
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "  '/최종수정자 ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSTIMESTAMP "                '/최종수정일자
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "  '/처방일자
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & lbl처방SEQ & " "                                     '/처방번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & lbl병록번호 & "' "                                   '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & lbl처방코드 & "' "                                   '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                     '/원처방 살아있는 처방
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                     '/처방History 번호
            If RunSQL(gstrQuy) = False Then Exit Function

            '/실시처방: TMPRSCEXCN
            gstrQuy = "UPDATE TMPRSCEXCN SET "
            gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '560', "                      '/560.임시결과 이상
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "  '/최종수정자 ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSTIMESTAMP "                '/최종수정일자
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "  '/처방일자
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & lbl처방SEQ & " "                                     '/처방번호
            gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & lbl병록번호 & "' "                                   '/병록번호
            gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & lbl처방코드 & "' "                                   '/처방코드
            gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                     '/실시처방 살아있는 처방
            If RunSQL(gstrQuy) = False Then Exit Function
  
        Case Else
            MsgBox "최종결과 처리에 대한 HIS 정보가 없습니다!", vbCritical, "경고"
    End Select
    
    FUNC_HIS_RST_UPDATE = True
End Function

Public Function FUNC_HIS_ORDER1_VIEW(argOrderCode As String) As Boolean '/미전송 처방조회(서울시립서북병원)
    FUNC_HIS_ORDER1_VIEW = False
    
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "       A.PID AS CHRTNO, "                                                            '/병록번호
    gstrQuy = gstrQuy & vbCrLf & "       DECODE(A.PRSC_OCRR_DVCD, 'I', '입원', 'O', '외래', '기타') AS IO_SECTION, "   '/외래/입원구분(I:입원, O:외래)
    gstrQuy = gstrQuy & vbCrLf & "       C.DEPT_ENGL_ABNM AS DETPCD, "                                                 '/진료과 약칭(영문)
    gstrQuy = gstrQuy & vbCrLf & "       B.PT_NM AS PATNM, "                                                           '/수진자명
    gstrQuy = gstrQuy & vbCrLf & "       B.SEX_CD AS SEX, "                                                            '/성별
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_1 AS JUMIN1, "                                                      '/주민번호1
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_2 AS JUMIN2, "                                                      '/주민번호2
    gstrQuy = gstrQuy & vbCrLf & "       fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "          '/HIS 나이계산 함수
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE AS ORDDATE, "                                                     '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NO AS ORDSEQ, "                                                        '/처방번호
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_CD AS ORDCD, "                                                         '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NM AS ORDNM, "                                                         '/처방명
    gstrQuy = gstrQuy & vbCrLf & "       A.DLVR_MATR, "                                                                '/전달사항
    gstrQuy = gstrQuy & vbCrLf & "       A.SUPT_DEPT_DLVR_MATR, "                                                      '/지원부서 전달사항
    gstrQuy = gstrQuy & vbCrLf & "       A.CNDT_PRSC_STAT_CD "                                                         '/실시처방 처방진행상태Flag
    gstrQuy = gstrQuy & vbCrLf & "  FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C "                                     '/VPRSCINFN(처방조회 VIEW), TPAPTMASTN(환자마스터), TZDEPTMSTN(부서마스터)
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.MDCR_DPMT_CD        = C.DEPT_CD "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "            '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE           = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "            '/검사실 접수일자
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_VALD_YN        = 'Y' "                                                 '/원처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/실시처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_HSTR_CD        = 'O' "                                                 '/처방History 번호
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_STAT_CD   = '440' "                                               '/실시처방 처방진행상태Flag(440.접수)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE          <> '00000000' "                                          '/과내접수일자(미접수:00000000, 접수:유효일자)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD            IN (" & argOrderCode & ") "                              '/처방코드
    
    FUNC_HIS_ORDER1_VIEW = True
End Function

Public Function FUNC_HIS_ORDER2_VIEW(argOrderCode As String) As Boolean '/기전송 처방조회(서울시립서북병원)
    FUNC_HIS_ORDER2_VIEW = False
    
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "       A.PID AS CHRTNO, "                                                             '/병록번호
    gstrQuy = gstrQuy & vbCrLf & "       DECODE(A.PRSC_OCRR_DVCD, 'I', '입원', 'O', '외래', '기타') AS IO_SECTION, "    '/외래/입원구분(I:입원, O:외래)
    gstrQuy = gstrQuy & vbCrLf & "       C.DEPT_ENGL_ABNM AS DETPCD, "                                                  '/진료과 약칭(영문)
    gstrQuy = gstrQuy & vbCrLf & "       B.PT_NM AS PATNM, "                                                            '/수진자명
    gstrQuy = gstrQuy & vbCrLf & "       B.SEX_CD AS SEX, "                                                             '/성별
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_1 AS JUMIN1, "                                                       '/주민번호1
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_2 AS JUMIN2, "                                                       '/주민번호2
    gstrQuy = gstrQuy & vbCrLf & "       fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "           '/HIS 나이계산 함수
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE AS ORDDATE, "                                                      '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NO AS ORDSEQ, "                                                         '/처방번호
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_CD AS ORDCD, "                                                          '/처방코드
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NM AS ORDNM, "                                                          '/처방명
    gstrQuy = gstrQuy & vbCrLf & "       A.DLVR_MATR, "                                                                 '/전달사항
    gstrQuy = gstrQuy & vbCrLf & "       A.SUPT_DEPT_DLVR_MATR, "                                                       '/지원부서 전달사항
    gstrQuy = gstrQuy & vbCrLf & "       D.EXAMDATE, "                                                                  '/최종결과일자
    gstrQuy = gstrQuy & vbCrLf & "       A.CNDT_PRSC_STAT_CD "                                                          '/실시처방 처방진행상태Flag
    gstrQuy = gstrQuy & vbCrLf & "  FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C, MM_EMR_RES D "                        '/VPRSCINFN(처방조회 VIEW), TPAPTMASTN(환자마스터), TZDEPTMSTN(부서마스터)
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.MDCR_DPMT_CD        = C.DEPT_CD "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = D.ORDDATE "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_NO             = D.ORDSEQ "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "            '/처방일자
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE           = '" & Format(CDate(dtp접수일자.Value), "YYYYMMDD") & "' "            '/검사실 접수일자
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_VALD_YN        = 'Y' "                                                 '/원처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/실시처방 살아있는 처방
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_HSTR_CD        = 'O' "                                                 '/처방History 번호
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_STAT_CD  >  '440' "                                               '/실시처방 처방진행상태Flag(560.임시결과 이상)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE          <> '00000000' "                                          '/과내접수일자(미접수:00000000, 접수:유효일자)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD            IN (" & argOrderCode & ") "                              '/처방코드
    
    FUNC_HIS_ORDER2_VIEW = True
End Function

