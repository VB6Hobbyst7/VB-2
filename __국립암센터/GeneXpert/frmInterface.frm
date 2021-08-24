VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmInterface 
   Caption         =   " GENEXPERT Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10680
   ScaleWidth      =   15030
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame FrmUseControl 
      Caption         =   "UseControl"
      Height          =   3195
      Left            =   3150
      TabIndex        =   49
      Top             =   4830
      Visible         =   0   'False
      Width           =   8385
      Begin VB.TextBox Text_Today 
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   6330
         TabIndex        =   74
         Text            =   "2002/02/18"
         Top             =   1170
         Width           =   1515
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "test"
         Height          =   555
         Left            =   4710
         TabIndex        =   73
         Top             =   2010
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtTest 
         Height          =   855
         Left            =   210
         TabIndex        =   72
         Top             =   1590
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   360
         TabIndex        =   63
         Top             =   1020
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.OptionButton optGBN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bar"
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   61
         Top             =   480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optGBN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seq"
         Height          =   255
         Index           =   1
         Left            =   2310
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   1020
         Top             =   330
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   150
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin FPSpread.vaSpread vasModuleCnt 
         Height          =   495
         Left            =   3990
         TabIndex        =   62
         Top             =   330
         Width           =   3915
         _Version        =   393216
         _ExtentX        =   6906
         _ExtentY        =   873
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   1
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin MSComCtl2.DTPicker Text_Today2 
         Height          =   345
         Left            =   6660
         TabIndex        =   75
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127598593
         CurrentDate     =   40248
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5670
         TabIndex        =   76
         Top             =   2700
         Width           =   840
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10305
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14286
            MinWidth        =   14286
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2020-12-30"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   3245
            MinWidth        =   3245
            TextSave        =   "오후 2:55"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame SSPanel1 
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5340
         Picture         =   "frmInterface.frx":08C7
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   77
         Top             =   420
         Width           =   315
      End
      Begin IF_GENEXPERT.MDButton cmdSend 
         Height          =   585
         Left            =   12450
         TabIndex        =   64
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "선택전송"
      End
      Begin IF_GENEXPERT.MDButton cmdReset 
         Height          =   585
         Left            =   11250
         TabIndex        =   65
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "화면정리"
      End
      Begin IF_GENEXPERT.MDButton cmdCall 
         Height          =   585
         Left            =   10050
         TabIndex        =   66
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "결과조회"
      End
      Begin IF_GENEXPERT.MDButton cmdExit 
         Height          =   585
         Left            =   13650
         TabIndex        =   67
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "종 료"
      End
      Begin MSComCtl2.DTPicker dtpReceDate 
         Height          =   345
         Left            =   8070
         TabIndex        =   69
         Top             =   390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127598593
         CurrentDate     =   40248
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5760
         TabIndex        =   78
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   360
         TabIndex        =   71
         Top             =   600
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7080
         TabIndex        =   70
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "GENEXPERT INTERFACE"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   210
         TabIndex        =   68
         Top             =   150
         Width           =   4035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9375
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   14925
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   660
         TabIndex        =   46
         Top             =   300
         Width           =   195
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   9015
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   7815
         _Version        =   393216
         _ExtentX        =   13785
         _ExtentY        =   15901
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         ShadowDark      =   0
         SpreadDesigner  =   "frmInterface.frx":0E51
         UserResize      =   2
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   9015
         Left            =   8040
         TabIndex        =   48
         Top             =   240
         Width           =   6765
         _Version        =   393216
         _ExtentX        =   11933
         _ExtentY        =   15901
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   8
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":187D
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4455
      Left            =   210
      TabIndex        =   3
      Top             =   4740
      Visible         =   0   'False
      Width           =   13095
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   3720
         TabIndex        =   31
         Top             =   2790
         Width           =   9285
         Begin VB.TextBox txtEquipID 
            Height          =   345
            Left            =   3600
            TabIndex        =   42
            Text            =   "10"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rack Pos"
            Height          =   375
            Left            =   7560
            TabIndex        =   41
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command10 
            Caption         =   "결과입력"
            Height          =   375
            Left            =   5880
            TabIndex        =   40
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEquipCode 
            Height          =   345
            Left            =   1710
            TabIndex        =   39
            Text            =   "0ADVI120"
            Top             =   1125
            Width           =   1875
         End
         Begin VB.CommandButton Command9 
            Caption         =   "장비ID조회"
            Height          =   375
            Left            =   60
            TabIndex        =   38
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command8 
            Caption         =   "미검사상세목록"
            Height          =   375
            Left            =   5010
            TabIndex        =   37
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command7 
            Caption         =   "미검사목록"
            Height          =   375
            Left            =   3360
            TabIndex        =   36
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Caption         =   "검사상세목록"
            Height          =   375
            Left            =   1710
            TabIndex        =   35
            Top             =   690
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   6660
            TabIndex        =   34
            Text            =   "05111000003"
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton Command5 
            Caption         =   "검사목록"
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command4 
            Caption         =   "서버시간"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   1920
            TabIndex        =   44
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   3150
            TabIndex        =   43
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   210
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3360
         Width           =   945
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   28
         Top             =   1950
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   240
         TabIndex        =   26
         Top             =   1380
         Width           =   3045
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   60
         TabIndex        =   19
         Top             =   3780
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txtEnd 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1950
            TabIndex        =   22
            Top             =   180
            Width           =   885
         End
         Begin VB.TextBox txtStart 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   630
            TabIndex        =   21
            Top             =   180
            Width           =   885
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   20
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " - "
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   23
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Top             =   1875
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   10260
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   17
         Top             =   1950
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   1260
         Picture         =   "frmInterface.frx":1DF1
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   2010
         Picture         =   "frmInterface.frx":1F20
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   240
         TabIndex        =   11
         Top             =   2850
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton cmdResSave 
         Caption         =   "결과저장"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5970
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 작성"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   7
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   2355
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   1710
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   10740
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":2052
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   3450
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":651B
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1125
         Left            =   7110
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":6762
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   5295
         TabIndex        =   25
         Top             =   240
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":69A9
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   8925
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":6BF0
      End
      Begin VB.Label lblMT 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9750
         TabIndex        =   45
         Top             =   2370
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.Frame FrmTempBox 
      Caption         =   "TempBox"
      Height          =   2205
      Left            =   1080
      TabIndex        =   50
      Top             =   7710
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CommandButton cmdQC 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6030
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command14 
         Caption         =   "사용자변경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1650
         TabIndex        =   57
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdResCall 
         Caption         =   "QC 결과전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   56
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   435
         Left            =   90
         TabIndex        =   55
         Top             =   1050
         Width           =   2325
      End
      Begin VB.CommandButton Command_setup 
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2310
         TabIndex        =   54
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_close 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3420
         TabIndex        =   53
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_Config 
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colSeqNo = 3
Const colReceno = 4
Const colRack = 5
Const colPos = 6
Const colPID = 7
Const colPName = 8
Const colPSex = 9
Const colPAge = 10
Const colPJumin = 11
Const colState = 12

Const colOrd = 13
Const colRes = 14
Const colDate = 15
Const colTime = 16
Const colTestType = 17

Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colRCheck = 6

'2004/10/21 이상은
'Const colRefLow = 7
Const colResult1 = 7

Const colRefHigh = 8

Dim gRow As Long

Dim gsBarCode As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const rs  As String = ""
Const GS  As String = ""

Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================

Dim gSelExam        As String

Function Advia_IDSet(asID As String) As String
    Advia_IDSet = "000" & asID
End Function

Function CX_Init() As String
'    Dim lsData As String
'
'    gMT = "0"
'    gErrState = 0
'
'    lsData = "[00,801,01]"
'    lsData = lsData & ASTM_CSum(lsData) & chrCR & chrLF
'
'    gComState = 0
'
'    gPreMsg = lsData
'    MSComm1.Output = lsData
'    Timer1.Enabled = True
'    SaveData "[Tx]" & lsData
End Function

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub cmdCall_Click()
Dim iRow As Long

    ClearSpread vasID
    ClearSpread vasRes
    
    SQL = "select distinct levelname, '', '', '0', '0', examtime, '', '', '', 'F' " & vbCrLf & _
          "from qc_res " & vbCrLf & _
          "where equipno  = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and examdate = '" & Format(CDate(dtpReceDate.Value), "yyyymmdd") & "' "
    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), count(*), max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(dtpReceDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno, pid, pname, page, psex, jumin, sendflag "
    SQL = SQL & vbCrLf & " Union " & vbCrLf
    SQL = SQL & vbCrLf & _
          "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), '0',  max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(dtpReceDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag not in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno,  pid, pname, page, psex, jumin, sendflag " & vbCrLf & _
          "order by diskno,posno"
    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
    
'    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue, max(recedate)" & _
'          "from pat_res " & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "group by barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue " & vbCrLf & _
'          "order by diskno,posno"
'    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    vasSort vasID, colRack, colPos
    
    For iRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, iRow, colState))
        Case "B", "C"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasID, "완료", iRow, colState
'        Case "C"
'            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
'            SetForeColor vasID, iRow, iRow, colState, colState, 255, 0, 0
'            SetText vasID, "완료(Alarm)", iRow, colState
        Case "O"
            SetText vasID, "오더", iRow, colState
         Case "A"
            SetText vasID, "결과", iRow, colState
        End Select
    Next iRow
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    Dim lsPID As String
    Dim lsReceNo1 As String
    Dim lsReceNo2 As String
    
    Dim sStart As String
    Dim send As String
    
    sStart = Trim(txtStart.Text)
    send = Trim(txtEnd.Text)
    
    If sStart <> "" And send <> "" Then
        For lRow = sStart To send
            lsPID = Trim(GetText(vasID, lRow, 5))
            lsReceNo1 = Trim(GetText(vasID, lRow, 11))
            lsReceNo2 = Trim(GetText(vasID, lRow, 12))
            SQL = "Delete from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and pid = '" & lsPID & "' " & vbCrLf & _
                  "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                  "  and receno1 = '" & lsReceNo2 & "' "
            res = SendQuery(gLocal, SQL)
            
            DeleteRow vasID, lRow, lRow
        Next lRow
    Else
        lRow = 1
        Do While lRow <= vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1
            If vasID.Value = 1 Then
                lsPID = Trim(GetText(vasID, lRow, 5))
                lsReceNo1 = Trim(GetText(vasID, lRow, 11))
                lsReceNo2 = Trim(GetText(vasID, lRow, 12))
                SQL = "Delete from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and pid = '" & lsPID & "' " & vbCrLf & _
                      "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                      "  and receno1 = '" & lsReceNo2 & "' "
                res = SendQuery(gLocal, SQL)
                
                DeleteRow vasID, lRow, lRow
            Else
                lRow = lRow + 1
            End If
        Loop
    End If
    
    MsgBox "삭제 완료"
    chkAll.Value = 0
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdOrder_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
    
Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdResCall_Click()
'    frmResult.Show 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
'    If chkAll.Value = 1 Then
'            For i = 1 To vasID.DataRowCnt
'                vasID.Row = i
'                vasID.Col = 1
'
'                If vasID.Value = 1 Then
'                    DeleteRow vasID, i, i
'                    i = i - 1
'                End If
'            Next i
'
'            chkAll.Value = 0
'    Else
'        vasID.Row = 1
'        vasID.Row2 = vasID.MaxRows
'        vasID.Col = 1
'        vasID.Col2 = vasID.MaxCols
'        vasID.BlockMode = True
'        vasID.BackColor = RGB(255, 255, 255)
'        vasID.Action = 3
'        vasID.BlockMode = False
'    End If
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    ClearSpread vasID
    ClearSpread vasRes
    
    Text_Today = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
    If MSComm1.PortOpen = True Then CX_Init
    
End Sub

Private Sub cmdResSave_Click()
'    Proc_Result txtBarcode
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            res = Insert_Data(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, lRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdTest_Click()
    Dim lsChar As String
    Dim i As Long
    
    For i = 1 To Len(txtTest.Text)
        lsChar = Mid(txtTest.Text, i, 1)
        Select Case lsChar
        Case chrSTX
            txtData.Text = lsChar
        Case chrETX
            txtData.Text = txtData.Text & "COL            Straw" & Chr(13) & Chr(10)
            txtData.Text = txtData.Text & "TUR            Clear" & Chr(13) & Chr(10)
            txtData.Text = txtData.Text & lsChar
            SaveData "[RX]" & txtData.Text
            AX4030 txtData.Text
        Case Else
            txtData.Text = txtData.Text & lsChar
        End Select
    Next
    
'    For i = 1 To Len(txtTest.Text)
'
'
'        lsChar = Mid(txtTest.Text, i, 1)
'
'        Select Case lsChar
'        Case chrENQ
'            SaveData "[RX]" & lsChar
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'        Case chrSTX
'            txtData.Text = lsChar
'        Case chrETB
'            txtData.Text = txtData.Text & lsChar
'            CliniTekAll Mid(txtData.Text, 3)
'        Case chrETX
'            txtData.Text = txtData.Text & lsChar
'            CliniTekAll Mid(txtData.Text, 3)
'        Case chrLF
'            txtData.Text = txtData.Text & lsChar
'            SaveData "[RX]" & txtData.Text
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'        Case chrEOT
'        Case Else
'            txtData.Text = txtData.Text & lsChar
'        End Select
'    Next
    
    txtTest.Text = ""
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

Private Sub Command_config_Click()
    frmConfig.Show 1
End Sub


Private Sub Command_setup_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub Command1_Click()
'    Hitachi747 Mid(Text2.Text, 2)
End Sub

Private Sub Command13_Click()
    Dim i As Integer
    
    SQL = "select item_code, item_name, m_stype_code, disp_seq, m_item_code from tbl_item"
    res = db_select_Vas(gLocal_1, SQL, vaSpread1)
    
    SQL = "delete from equipexam"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vaSpread1.DataRowCnt
    
        SQL = "insert into equipexam(equipno, examcode, equipcode, examname, examtype, seqno, resprec, examflag) " & vbCrLf & _
              "values('C064','" & Trim(GetText(vaSpread1, i, 1)) & "','" & Trim(GetText(vaSpread1, i, 5)) & "','" & Trim(GetText(vaSpread1, i, 2)) & "','" & Trim(GetText(vaSpread1, i, 3)) & "','" & Trim(GetText(vaSpread1, i, 4)) & "','1','1')"
        res = SendQuery(gLocal, SQL)
    Next
    
End Sub

Private Sub Command14_Click()
'    frmUserChange.Show 0
    
End Sub

Private Sub Command15_Click()
    
    Online_XML gXml_S07, "10010700001"
'    vasID.MaxRows = 1
'    SetText vasID, "10010700001", 1, colBarCode
'
'    Get_Sample_Info 1
'
End Sub

Private Sub Command3_Click()
    SQL = "CREATE INDEX resindex1 ON pat_res (examdate,equipno,barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex1 created"
    Else
        MsgBox "resindex1 failed"
    End If
    SQL = "CREATE INDEX resindex2 ON pat_res (examdate,equipno,barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex2 created"
    Else
        MsgBox "resindex2 failed"
    End If
    
    SQL = "CREATE INDEX resindex3 ON pat_res (barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex3 created"
    Else
        MsgBox "resindex3 failed"
    End If
    
    SQL = "CREATE INDEX resindex4 ON pat_res (barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex4 created"
    Else
        MsgBox "resindex4 failed"
    End If
End Sub

Private Sub Form_Load()
    Dim sDate As String
            
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    cmdReset_Click
    
    GetSetup
    
'    MSComm1.CommPort = gSetup.gPort
'    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
'
'    If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    Text_Today = Format(Date, "yyyy/mm/dd")

    GetExamCode
    
    lblUser.Caption = gIFUser
    
    sDate = Format(DateAdd("y", CDate(dtpReceDate.Value), -50), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    dtpReceDate = Format(Date, "yyyy-mm-dd")
    
    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    
    'Call ExamCount
    lblUser.Caption = gIFUser
    
    wSck.LocalPort = CInt(gServerPort)
    wSck.Listen
    
    lblStatus.Caption = ": TCP " & gServerPort & " 포트로 연결합니다"
    
    
End Sub

'Sub ExamCount()
'    SQL = "SELECT COUNT(DISKNO)"
'    SQL = SQL & vbCrLf & "  FROM PAT_RES"
'    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
'    SQL = SQL & vbCrLf & "  AND DISKNO = '1'"
'    res = db_select_Col(gLocal, SQL)
'
'    If gReadBuf(0) = "" Then
'        SetText vasModuleCnt, "0", 1, 1
'    Else
'        SetText vasModuleCnt, gReadBuf(0), 1, 1
'    End If
'
'    SQL = "SELECT COUNT(DISKNO)"
'    SQL = SQL & vbCrLf & "  FROM PAT_RES"
'    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
'    SQL = SQL & vbCrLf & "  AND DISKNO = '2'"
'    res = db_select_Col(gLocal, SQL)
'
'    If gReadBuf(0) = "" Then
'        SetText vasModuleCnt, "0", 1, 2
'    Else
'        SetText vasModuleCnt, gReadBuf(0), 1, 2
'    End If
'
'    SQL = "SELECT COUNT(DISKNO)"
'    SQL = SQL & vbCrLf & "  FROM PAT_RES"
'    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
'    SQL = SQL & vbCrLf & "  AND DISKNO = '3'"
'    res = db_select_Col(gLocal, SQL)
'
'    If gReadBuf(0) = "" Then
'        SetText vasModuleCnt, "0", 1, 3
'    Else
'        SetText vasModuleCnt, gReadBuf(0), 1, 3
'    End If
'
'    SQL = "SELECT COUNT(DISKNO)"
'    SQL = SQL & vbCrLf & "  FROM PAT_RES"
'    SQL = SQL & vbCrLf & " WHERE COUNTYN IS NULL "
'    SQL = SQL & vbCrLf & "  AND DISKNO = '4'"
'    res = db_select_Col(gLocal, SQL)
'
'    If gReadBuf(0) = "" Then
'        SetText vasModuleCnt, "0", 1, 4
'    Else
'        SetText vasModuleCnt, gReadBuf(0), 1, 4
'    End If
'
'End Sub


'-- 일련번호
'''Function Get_Sample_Info(ByVal asRow As Long) As Integer
'''Dim lsBarcode As String
'''Dim lsReceDate As String
'''Dim lsExamPart As String
'''
'''Dim lsReceno As String
'''Dim lsPID As String
''''Dim lsReceno As String
'''Dim sRes As String
'''
'''
'''    Get_Sample_Info = -1
'''
'''    '샘플 환자 정보 가져오기
'''    If Mid(Right(Format(Trim(GetText(vasID, asRow, colReceno)), "0000"), 4), 1, 1) = "3" Then
'''        lsExamPart = "L80"
'''    Else
'''        lsExamPart = "L61"
'''    End If
'''
'''
'''    lsReceno = Trim(GetText(vasID, asRow, colReceno))    '샘플 바코드 번호
'''    lsReceDate = Format(dtpReceDate.Value, "yyyymmdd")
''''    lsExamPart = "L80"
'''
''''    If Trim(lsbarcode) = "" Then: Exit Function
'''    sRes = Online_TLA(gXml_S18, lsReceDate, lsReceno, lsExamPart)
''''    If sRes = 1 Then
'''    SetText vasID, gPat_Info_Select.SPC_NO, asRow, colBarCode
'''    SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
'''    SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName
'''    SetText vasID, gPat_Info_Select.SEX, asRow, colPSex
'''    SetText vasID, gPat_Info_Select.AGE, asRow, colPAge
'''    SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
'''    SetText vasID, Format(gPat_Info_Select.ACPT_DTETM, "yyyymmdd"), asRow, colDate
'''    SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno
'''
''''        vasID.RowHeight(asRow) = 20
'''
'''        Get_Sample_Info = 1
''''    End If
'''End Function

'-- 바코드
Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsBarcode As String
Dim lsPID As String
Dim lsReceNo As String
Dim sRes As String
Dim lsReceDate As String
Dim lsExamPart As String

    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsBarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
'''    Save_Raw_Data "[lsBarcode]" & lsBarcode
    
    
'    If Trim(lsbarcode) = "" Then: Exit Function
    sRes = Online_XML(gXml_S03, lsBarcode)
'    If sRes = 1 Then
        SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
        SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName
        SetText vasID, gPat_Info_Select.SEX, asRow, colPSex
        SetText vasID, gPat_Info_Select.AGE, asRow, colPAge
        SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
        SetText vasID, gPat_Info_Select.ACPT_DTETM, asRow, colDate
        SetText vasID, Mid(gPat_Info_Select.ACPT_DTETM, 1, 10), asRow, colDate
        SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno

        vasID.RowHeight(asRow) = 12
        
        Get_Sample_Info = 1
'    End If


End Function

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    
    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    SQL = "Select equipcode, examcode, examname, reflow, refhigh " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 6)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 5
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub MDButton3_Click()

End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1

End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0

End Sub

Private Sub MSComm1_OnComm()

    Select Case MSComm1.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

            Buffer = MSComm1.Input
            Save_Raw_Data "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                intPhase = 2
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case ACK
                                'If strState = "Q" Then Call SendOrder
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
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
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
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
                                intPhase = 4
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
'                            Case vbLf
'                                intPhase = 4
'                                MSComm1.Output = ACK
'                                Save_Raw_Data "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                'Call EditRcvData
                                Call AX4030(strRecvData)
'                                If strState = "Q" Then
'                                    intSndPhase = 1
'                                    intFrameNo = 1
'                                    MSComm1.Output = ENQ
'                                    Save_Raw_Data "[Tx]" & ENQ
'                                End If
                                intPhase = 1
                        End Select
                End Select
            Next i
            
        End Select


End Sub


Private Sub SendWSckData(ByVal pSendData As Variant)

    '-- 전송
    wSck.SendData pSendData
    
    '-- 로그기록
    Call SaveData("[Tx]" & pSendData)

End Sub


Sub RcvSocketData(ByVal lsData As String)

        Dim Buffer      As Variant
        Dim BufChar     As String
        Dim lngBufLen   As Long
        Dim i           As Long

        Buffer = lsData
        lngBufLen = Len(Buffer)
        
        Debug.Print Buffer
        
        For i = 1 To lngBufLen
            BufChar = Mid$(Buffer, i, 1)

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
                        intPhase = 1
                        Call GENEXPERT
                        
                End Select
        End Select
        Next i
            

End Sub


Sub AX4030(asVar As Variant)
'Sub AX4030(asVar As String)
    Dim ResultTbl(1 To 100) As String
    Dim TablePtr As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim icnt As Integer
    Dim lsID As String
    Dim lsSeq As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsBarcode As String
    Dim lsResult As String
    Dim lsExamDate As String
    Dim lsGubun As Boolean
    Dim lsRow As Integer
    Dim lsResRow As Integer
    Dim lsRack As String
    Dim lsPos As String
    Dim lsEquipRes As String
    
    Dim intCnt As Integer
    Dim strRcvBuf As String
    Dim strType As String
    Dim ii As Integer
    Dim lsSelExam   As String
                
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        Select Case strType
        Case "O"
            'bar
            If optGBN(0).Value = True Then
                
                lsBarcode = Trim(mGetP(mGetP(strRecvData(intCnt), 3, "|"), 1, "^"))
                If InStr(lsBarcode, "----") > 0 Then
'                    Exit Sub
                End If
                lsRow = -1
                For i = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, i, colBarCode)) = lsBarcode Then
                        lsRow = i
                        Exit For
                    End If
                Next
                If lsRow = -1 Then
                    For i = 1 To vasID.DataRowCnt
                        If Trim(GetText(vasID, i, colBarCode)) = "" Then
                            lsRow = i
                            Exit For
                        End If
                    Next
                End If
            
                If lsRow = -1 Then
                    lsRow = vasID.DataRowCnt + 1
                    If vasID.MaxRows < lsRow Then
                        vasID.MaxRows = lsRow
                    End If
                End If
                
                gRow = lsRow
                
                SetText vasID, lsSeq, gRow, colRack
                SetText vasID, lsSeq, gRow, colReceno
                
                SetText vasID, lsBarcode, gRow, colBarCode
            
                gOrderExam = ""
                Get_Sample_Info gRow
            
                '**************************************************
                res = Online_XML(gXml_S07, Trim(lsBarcode))
                
                ClearSpread vasTemp
                
                
                lsSelExam = ""
                
                gSelExam = ""
                
                For ii = 0 To UBound(gExam_Select)
                    vasTemp.SetText 1, ii + 1, gExam_Select(ii).TST_CD
                    If lsSelExam = "" Then
                        lsSelExam = "'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
                    Else
                        lsSelExam = lsSelExam & ",'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
                    End If
                Next ii
                
                gSelExam = lsSelExam
                '**************************************************
            
                If vasRes.MaxRows < lsResRow Then
                    vasRes.MaxRows = lsResRow
                End If
            
            'seq
            Else
                lsSeq = mGetP(strRecvData(intCnt), 4, "|")
                
                lsRow = -1
                For i = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, i, colRack)) = lsSeq Then
                        lsRow = i
                        Exit For
                    End If
                Next
                If lsRow = -1 Then
                    For i = 1 To vasID.DataRowCnt
                        If Trim(GetText(vasID, i, colRack)) = "" Then
                            lsRow = i
                            Exit For
                        End If
                    Next
                End If
                
                If lsRow = -1 Then
                    lsRow = vasID.DataRowCnt + 1
                    If vasID.MaxRows < lsRow Then
                        vasID.MaxRows = lsRow
                    End If
                End If
                
                gRow = lsRow
                
                SetText vasID, lsSeq, gRow, colRack
                SetText vasID, lsSeq, gRow, colReceno
                
                gOrderExam = ""
                Get_Sample_Info gRow
                
                If Trim(GetText(vasID, gRow, colBarCode)) = "" Then
                    lsBarcode = Format(Date, "yymmdd") & "-" & Format(lsSeq, "00000")
                    SetText vasID, lsBarcode, gRow, colBarCode
                End If
                
                'lsResRow = X - 6
                If vasRes.MaxRows < lsResRow Then
                    vasRes.MaxRows = lsResRow
                End If
                
            End If
            
        Case "R"
            lsResRow = lsResRow + 1
            lsEquipCode = Trim(mGetP(mGetP(strRecvData(intCnt), 3, "|"), 4, "^"))
            
            If lsEquipCode = "PH" Or lsEquipCode = "S.G." Or lsEquipCode = "COLOR" Or lsEquipCode = "LEU" Then
                lsResult = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 2, "^"))
                lsEquipRes = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 2, "^"))
                
                '4R|9|^^^LEU |-     ^|Leu/uL ||||F|||201202081157
                '4R|9|^^^LEU |^75|Leu/uL ||*||F|||201202081157
                If lsEquipCode = "LEU" Then
                    Select Case lsResult
                    Case "25"
                        lsResult = "+/-"
                    Case "75"
                        lsResult = "1+"
                    Case "250"
                        lsResult = "2+"
                    Case "500"
                        lsResult = "3+"
                    Case Else
                        lsResult = lsResult
                    End Select
                    
                    If lsResult = "" Then
                        lsResult = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 1, "^"))
                        lsEquipRes = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 1, "^"))
                    End If
                ElseIf lsEquipCode = "S.G." Then
                    If UCase(lsResult) = "OVER" Then
                        lsResult = "1.050"
                    End If
                Else
                    If lsResult = "0.2" Or lsResult = "1.0" Then
                        lsResult = "+/-"
                    ElseIf lsResult = "2.0" Then
                        lsResult = "1+"
                    ElseIf lsResult = "4.0" Then
                        lsResult = "2+"
                    ElseIf lsResult = ">=8.0" Then
                        lsResult = "3+"
                    End If
                End If
            Else
                lsResult = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 1, "^"))
                lsEquipRes = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 1, "^"))
                If lsEquipCode = "SG" Then
                    lsResult = Replace(lsResult, "=", "")
                    lsResult = Replace(lsResult, "<", "")
                    lsResult = Replace(lsResult, ">", "")
                End If
            
            End If
            
            If lsEquipCode = "COLOR" Then
                lsResult = "Straw"
            End If
            
            If lsEquipCode = "TURB" Then
                Select Case lsResult
                Case "-"
                    lsResult = "Clear"
                Case Else
                    lsResult = "Cloudy"
                End Select
            End If
            
            
            If lsEquipCode = "URO" Then
                Select Case lsResult
                Case "NORMAL"
                    lsResult = "+/-"
                Case Else
                    lsResult = lsResult
                End Select
            '-- 추가
            ElseIf lsEquipCode = "BLD" Then
                Select Case lsResult
                Case "+-"
                    lsResult = "+/-"
                Case Else
                    lsResult = lsResult
                End Select
            Else
                Select Case lsResult
                Case "NEGATIVE"
                    lsResult = "-"
                Case "TRACE"
                    lsResult = "+/-"
                Case "POSITIVE"
                    lsResult = "1+"
                Case Else
                    lsResult = lsResult
                End Select
            End If
        
            If lsResult = "+-" Then
                lsResult = "+/-"
            End If
            
            
            'bar
            If optGBN(0).Value = True Then
                SQL = "Select ExamCode, ExamName From EquipExam " & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And  EquipCode = '" & Trim(lsEquipCode) & "'" & vbCrLf & _
                      "  And  ExamCode in (" & gSelExam & ") "
                res = db_select_Col(gLocal, SQL)
                
                If res = 1 And gReadBuf(0) <> "" Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                End If
                
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                
                SetText vasRes, lsEquipCode, lsResRow, colEquipCode
                SetText vasRes, lsExamCode, lsResRow, colExamCode
                SetText vasRes, lsExamName, lsResRow, colExamName
                SetText vasRes, lsResult, lsResRow, colResult
                SetText vasRes, lsEquipRes, lsResRow, colResult1                            '검사결과
                
                Save_Local_One_1 gRow, lsResRow, "A"
            
            Else
                SQL = "select examcode, examname from equipexam where equipcode = '" & lsEquipCode & "' and examcode in (" & gOrderExam & ")"
                res = db_select_Col(gLocal, SQL)
                
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                
                SetText vasRes, lsEquipCode, lsResRow, colEquipCode
                SetText vasRes, lsExamCode, lsResRow, colExamCode
                SetText vasRes, lsExamName, lsResRow, colExamName
                SetText vasRes, lsResult, lsResRow, colResult
                SetText vasRes, lsEquipRes, lsResRow, colResult1                            '검사결과
                
                Save_Local_One_1 gRow, lsResRow, "A"
            
            
            End If
            
        Case "L"
            SetText vasID, "결과", gRow, colState
            
            If MnTransAuto.Checked = True And Len(Trim(GetText(vasID, gRow, colBarCode))) = 11 Then
                res = Insert_Data(gRow)
                
                strState = ""
                
                If res = -1 Then
                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                    SetText vasID, "실패", gRow, colState
                Else
                   
                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                    SetText vasID, "완료", gRow, colState
                    
                    SQL = " Update pat_res Set " & vbCrLf & _
                          " sendflag = 'C' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                End If
                        
            End If
                
        End Select
    Next
    


End Sub

Sub GENEXPERT()
    Dim ResultTbl(1 To 100) As String
    Dim TablePtr As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim icnt As Integer
    Dim lsID As String
    Dim lsSeq As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsBarcode As String
    Dim lsResult As String
    Dim lsExamDate As String
    Dim lsGubun As Boolean
    Dim lsRow As Integer
    Dim lsResRow As Integer
    Dim lsRack As String
    Dim lsPos As String
    Dim lsEquipRes As String
    
    Dim intCnt As Integer
    Dim strRcvBuf As String
    Dim strType As String
    Dim ii As Integer
    Dim lsSelExam   As String
    Dim strModule   As String
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        Save_Raw_Data "[RCV]" & strRcvBuf

        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        
        Select Case strType
        Case "H"
            strModule = ""
        Case "O"
            lsBarcode = mGetP(strRcvBuf, 3, "|")
            If lsBarcode = "" Then
                Exit Sub
            End If
            
            lsRow = -1
            For i = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, i, colBarCode)) = lsBarcode Then
                    lsRow = i
                    Exit For
                End If
            Next
            If lsRow = -1 Then
                For i = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, i, colBarCode)) = "" Then
                        lsRow = i
                        Exit For
                    End If
                Next
            End If
            If lsRow = -1 Then
                lsRow = vasID.DataRowCnt + 1
                If vasID.MaxRows < lsRow Then
                    vasID.MaxRows = lsRow
                End If
            End If
            
            gRow = lsRow
            Call SetText(vasID, lsBarcode, gRow, colBarCode)
            
            gOrderExam = ""
            Call Get_Sample_Info(gRow)
            res = Online_XML(gXml_S07, Trim(lsBarcode))
            
            If res > 0 Then
                Call ClearSpread(vasTemp)
                
                lsSelExam = ""
                gSelExam = ""
                
                For ii = 0 To UBound(gExam_Select)
                    vasTemp.SetText 1, ii + 1, gExam_Select(ii).TST_CD
                    If lsSelExam = "" Then
                        lsSelExam = "'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
                    Else
                        lsSelExam = lsSelExam & ",'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
                    End If
                Next ii
                gSelExam = lsSelExam
            End If
        
            If vasRes.MaxRows < lsResRow Then
                vasRes.MaxRows = lsResRow
            End If
            
        Case "R"
            lsResult = ""
            lsEquipCode = mGetP(strRcvBuf, 3, "|")
            lsEquipRes = Trim(mGetP(mGetP(strRecvData(intCnt), 4, "|"), 1, "^"))
            strModule = mGetP(mGetP(strRcvBuf, 14, "|"), 3, "^")
            
            '모듈번호를 보기쉽게 변경한다.
            If strModule = "614414" Then
                strModule = "1"
                SetText vasID, strModule, gRow, colRack
            ElseIf strModule = "614415" Then
                strModule = "2"
                SetText vasID, strModule, gRow, colRack
            ElseIf strModule = "619205" Then
                strModule = "3"
                SetText vasID, strModule, gRow, colRack
            ElseIf strModule = "633147" Then
                strModule = "4"
                SetText vasID, strModule, gRow, colRack
            Else
                If strModule <> "" Then
                    SetText vasID, strModule, gRow, colRack
                End If
            End If
            
            If lsEquipRes = "NEG" Then
                lsResult = "NEGATIVE"
            ElseIf lsEquipRes = "NEG" Then
                lsResult = "POSITIVE"
            ElseIf lsEquipRes = "DETECTED" Then
                lsResult = "Detected"
            ElseIf lsEquipRes = "NOT DETECTED" Then
                lsResult = "Not Detected"
            ElseIf lsEquipRes = "NEGATIVE" Then
                lsResult = "Negative"
            ElseIf lsEquipRes = "POSITIVE" Then
                lsResult = "Positive"
            Else
                If lsEquipCode = "^Xpert MTB-RIF A^^3^Xpert MTB-RIF Assay G4^5^Rif Resistance^" Then
                    lsEquipRes = "Not Detected"
                    lsResult = lsEquipRes
                Else
                    lsResult = lsEquipRes
                End If
            End If
RST:
            
            If lsEquipRes <> "" Then
                SQL = ""
                SQL = SQL & "Select ExamCode, ExamName From EquipExam           " & vbCrLf
                SQL = SQL & " Where Equipno     = '" & gEquip & "'              " & vbCrLf
                SQL = SQL & "   And EquipCode   = '" & Trim(lsEquipCode) & "'   " & vbCrLf
                If gSelExam <> "" Then
                    If gSelExam = "'L2742'" Then
                        gSelExam = "'L2742_1','L2742_2'"
                    ElseIf gSelExam = "'L2743'" Then
                        gSelExam = "'L2743_1','L2743_2','L2743_3'"
                    ElseIf gSelExam = "'L2759'" Then
                        gSelExam = "'L2759_1','L2759_2','L2759_3','L2759_4','L2759_5'"
                    End If
                    SQL = SQL & "  And  ExamCode in (" & gSelExam & ") "
                End If
                
                res = db_select_Col(gLocal, SQL)
                
                If res = 1 And gReadBuf(0) <> "" Then
                    lsResRow = lsResRow + 1
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                        
                    SetText vasRes, lsEquipCode, lsResRow, colEquipCode
                    SetText vasRes, lsExamCode, lsResRow, colExamCode
                    SetText vasRes, lsExamName, lsResRow, colExamName
                    SetText vasRes, lsResult, lsResRow, colResult
                    SetText vasRes, lsEquipRes, lsResRow, colResult1                            '검사결과
                    
                    Save_Local_One_1 gRow, lsResRow, "A"
                End If
            End If
            
        Case "L"
            SetText vasID, "결과", gRow, colState
            
            If MnTransAuto.Checked = True And Len(Trim(GetText(vasID, gRow, colBarCode))) = 11 Then
                res = Insert_Data(gRow)
                
                strState = ""
                
                If res = -1 Then
                    Call SetForeColor(vasID, gRow, gRow, 1, colState, 255, 0, 0)
                    Call SetText(vasID, "실패", gRow, colState)
                Else
                   
                    Call SetBackColor(vasID, gRow, gRow, 1, colState, 202, 255, 112)
                    Call SetText(vasID, "완료", gRow, colState)
                    
                    SQL = ""
                    SQL = SQL & "Update pat_res Set                 " & vbCrLf
                    SQL = SQL & "   sendflag    = 'C'               " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gEquip & "'  " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "'"
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                End If
                        
            End If
                
        End Select
    Next
    


End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


Sub CliniTekAll(asVar As String)
    Dim CheckType(1 To 10) As String
    Dim CheckCnt As Integer
    Dim i As Long
    
    CheckCnt = 1
    
    For i = 1 To Len(asVar)
        If Mid(asVar, i, 1) = chrCR Then
            CheckType(CheckCnt) = CheckType(CheckCnt) & Mid(asVar, i, 1)
            Clinitec CheckType(CheckCnt)
            CheckCnt = CheckCnt + 1
            CheckType(CheckCnt) = ""
        ElseIf Mid(asVar, i, 1) = chrETB Or Mid(asVar, i, 1) = chrETX Then
            Exit For
        Else
            CheckType(CheckCnt) = CheckType(CheckCnt) & Mid(asVar, i, 1)
        End If
    Next
End Sub

Sub Clinitec(asData As String)
    Dim ResultTbl(1 To 20) As String
    Dim TablePtr As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim icnt As Integer
    Dim lsID As String
    Dim lsSeq As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    
    Dim lsResult As String
    Dim lsExamDate As String
    Dim lsGubun As Boolean
    Dim lsRow As Integer
    Dim lsResRow As Integer
    
    
    TablePtr = 1
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            If TablePtr > 20 Then
                Exit For
            End If
            ResultTbl(TablePtr) = ""
        ElseIf (Mid(asData, j, 1) = chrCR) Then
            Exit For
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
    
    Select Case ResultTbl(1)
    Case "H"
        gRow = -1
        
    Case "P"
        lsRow = -1
        If ResultTbl(2) = "2" Then
            Exit Sub
        End If
        lsSeq = Trim(ResultTbl(3))
        
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colRack)) = lsSeq Then
                lsRow = i
                Exit For
            End If
        Next
        If lsRow = -1 Then
            For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colRack)) = "" Then
                lsRow = i
                Exit For
            End If
        Next
        End If
        
        If lsRow = -1 Then
            lsRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < lsRow Then
                vasID.MaxRows = lsRow
            End If
        End If
        gRow = lsRow
        SetText vasID, lsSeq, gRow, colRack
        
    Case "R"
        lsResRow = vasRes.DataRowCnt + 1
        If vasRes.MaxRows < lsResRow Then
            vasRes.MaxRows = lsResRow
        End If
        
        lsEquipCode = ResultTbl(4)
        lsResult = ResultTbl(6)
        i = InStr(1, lsResult, "^")
        If i > 0 Then
            lsResult = Mid(lsResult, 1, i - 1)
        End If
        
        SQL = "select examcode, examname from equipexam where equipcode = '" & lsEquipCode & "'"
        res = db_select_Col(gLocal, SQL)
        
        lsExamCode = Trim(gReadBuf(0))
        lsExamName = Trim(gReadBuf(1))
        
        SetText vasRes, lsEquipCode, lsResRow, colEquipCode
        SetText vasRes, lsExamCode, lsResRow, colExamCode
        SetText vasRes, lsExamName, lsResRow, colExamName
        SetText vasRes, lsResult, lsResRow, colResult
        SetText vasRes, lsResult, lsResRow, colResult1                            '검사결과
        Save_Local_One_1 gRow, lsResRow, "A"
        
    Case "L"
    End Select
    
End Sub

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
    If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
'        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
                      Trim(GetText(vasID, asRow1, colBarcode)), _
                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
                      Trim(GetText(vasRes, asRow2, colResult)), _
                      Trim(GetText(vasRes, asRow2, colResult1))
        'Save_Local_QC Trim(Text_Today.Text) & " " & Trim(GetText(vasID, asRow1, colPID)), _
                      Trim(GetText(vasID, asRow1, colBarcode)), _
                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
                      Trim(GetText(vasRes, asRow2, colResult)), _
                      Trim(GetText(vasRes, asRow2, colResult1))
        Exit Function
    End If

    sCnt = ""
    If Trim(GetText(vasRes, asRow2, colEquipCode)) = "" Then Exit Function
    
    SQL = ""
    SQL = SQL & "select count(*)    " & vbCrLf
    SQL = SQL & "  from pat_res     " & vbCrLf
    SQL = SQL & " where examdate    = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "'  " & vbCrLf
    SQL = SQL & "   and equipno     = '" & gEquip & "'                                      " & vbCrLf
    SQL = SQL & "   and barcode     = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "'    " & vbCrLf
    SQL = SQL & "   and equipcode   = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "' " & vbCrLf
    SQL = SQL & "   and examcode    = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'  "
    
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
    
    If sCnt = "0" Then
        SQL = ""
        SQL = SQL & "INSERT INTO pat_res "
        SQL = SQL & "(examdate   , equipno  , barcode   , seqno     , diskno  , posno   " & vbCrLf
        SQL = SQL & " ,pid       , pname    , jumin     , page      , psex    , resdate " & vbCrLf
        SQL = SQL & " ,receno    , equipcode, examcode  , result    , result1 , sendflag" & vbCrLf
        SQL = SQL & " ,examname  , refflag  , refvalue  , panicvalue, recedate )        " & vbCrLf
        SQL = SQL & " VALUES                                                            " & vbCrLf
        SQL = SQL & "('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "'              " & vbCrLf
        SQL = SQL & ",'" & Trim(gEquip) & "'                                            " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colBarCode)) & "'                " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'                  " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colRack)) & "'                   " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPos)) & "'                    " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPID)) & "'                    " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPName)) & "'                  " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPJumin)) & "'                 " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPAge)) & "'                   " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colPSex)) & "'                   " & vbCrLf
        SQL = SQL & ",'" & sExamDate & "'                                               " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colReceno)) & "'                 " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'             " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'              " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colResult)) & "'                " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colResult1)) & "'               " & vbCrLf
        SQL = SQL & ",'" & asSend & "'                                                  " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colExamName)) & "'              " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "'                " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colOrd)) & "'                    " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colRes)) & "'                    " & vbCrLf
        SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDate)) & "')"

        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
        SQL = ""
        SQL = SQL & "Update pat_res Set " & vbCrLf
        SQL = SQL & "   diskno      = '" & Trim(GetText(vasID, asRow1, colRack)) & "'      " & vbCrLf
        SQL = SQL & ",  posno       = '" & Trim(GetText(vasID, asRow1, colPos)) & "'       " & vbCrLf
        SQL = SQL & ",  result      = '" & Trim(GetText(vasRes, asRow2, colResult)) & "'   " & vbCrLf
        SQL = SQL & ",  result1     = '" & Trim(GetText(vasRes, asRow2, colResult1)) & "'  " & vbCrLf
        SQL = SQL & ",  refflag     = '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "'   " & vbCrLf
        SQL = SQL & ",  refvalue    = '" & Trim(GetText(vasID, asRow1, colOrd)) & "'       " & vbCrLf
        SQL = SQL & ",  panicvalue  = '" & Trim(GetText(vasID, asRow1, colRes)) & "'       " & vbCrLf
        SQL = SQL & ",  resdate     = '" & sExamDate & "'                                  " & vbCrLf
        SQL = SQL & " Where examdate    = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "'  " & vbCrLf
        SQL = SQL & "   And equipno     = '" & gEquip & "'                                      " & vbCrLf
        SQL = SQL & "   And barcode     = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "'    " & vbCrLf
        SQL = SQL & "   And equipcode   = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "' " & vbCrLf
        SQL = SQL & "   And examcode    = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'  "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Insert_Data(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim sDpcd, sDate1, sSlip, sItem, sOitp, sWkno As String
    Dim sIDNo, sSmyr, sSmsn, sSms1 As String
    Dim tSmsn As String
    Dim lsExamCode, lsResult As String
    Dim lPanicLow, lPanicHigh As Currency
    Dim lDeltaLow, lDeltaHigh, lDeltaMeth, lDeltaGap
    Dim lsPanic, lsDelta As String
    Dim lsPreDate, lsPreResult As String
    Dim lsNState, lsWState As String
    Dim lStdVal
    Dim lTerm As Long
    Dim lsQCChk As String

    Dim iNone, iDP

    Dim sResDate As String
    Dim sRDate As String
    Dim sRTime As String

    Dim lsID As String

    Dim i, j As Long
    Dim lRow As Long
    Dim lsQCOn As String
    
    Dim sResult As String
    Dim sExamCode As String
    Dim sBarCode As String
    Dim sEquipCode As String
    Dim sResStr As String
    Dim sResRow As Long
    Dim sResCnt As String
    Dim sEquipRes As String
    Dim sParam As String
    Dim X As Integer
    Dim lsTransCode As String
    
    Dim strXpertCd      As String
    Dim strXpertVal     As String
    Dim strEquip        As String
    Insert_Data = -1
    
    strXpertCd = ""
    strXpertVal = ""
    lsQCOn = ""

    lRow = argSpcRow

    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasID, lRow, colBarCode))
    sBarCode = ""
    sEquipCode = ""
    sResult = ""
    sExamCode = ""
    lsTransCode = ""
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    ClearSpread vasTemp1

    iNone = 0
    iDP = 0

    Online_XML gXml_S07, lsID
    
    SQL = ""
    SQL = SQL & "Select equipcode, examcode, examname, result, result1, diskno " & vbCrLf
    SQL = SQL & "  from pat_res                     " & vbCrLf
    SQL = SQL & " where equipno = '" & gEquip & "'  " & vbCrLf
    SQL = SQL & "   and barcode = '" & lsID & "'    " & vbCrLf
    SQL = SQL & "   and result  <> '' "
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    On Error GoTo ErrHandle
    
    sParam = ""
    strXpertCd = ""
    strXpertVal = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" Then
            strXpertCd = mGetP(Trim(GetText(vasTemp, sResRow, 2)), 1, "_")
            If strXpertCd = "L2808" Then
                strXpertVal = strXpertVal & "- 결과 : " & Trim(GetText(vasTemp, sResRow, 4)) & vbCrLf
            Else
                strXpertVal = strXpertVal & Trim(GetText(vasTemp, sResRow, 3)) & " : " & Trim(GetText(vasTemp, sResRow, 4)) & vbCrLf
            End If
            strEquip = Trim(GetText(vasTemp, sResRow, 6))
            If Trim(lsTransCode) = "" Then
                lsTransCode = "'" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
            Else
                lsTransCode = lsTransCode & ", '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
            End If
            
        End If
    Next
    
    If strXpertCd <> "" Then
        If strXpertCd = "L2743" Then
            strXpertVal = strXpertVal & vbCrLf
            strXpertVal = strXpertVal & "<결과해석>" & vbCrLf
            strXpertVal = strXpertVal & "1) Toxin B (+), Binary toxin (-), tcdC (-) : Toxin B를 분비하는 일반 C.diffcile 균주" & vbCrLf
            strXpertVal = strXpertVal & "2) Toxin B (+), Binary toxin 또는 tcdC 둘 중 하나 (+) : Toxin B를 분비하는 일반 C.diffcile 균주" & vbCrLf
            strXpertVal = strXpertVal & "   - Binary toxin (+) : toxin의 활성을 촉진하므로 적극적인 치료가 요구됨" & vbCrLf
            strXpertVal = strXpertVal & "   - tcdC (+) : toxin B의 분비가 약 23배 증가됨이 보고됨" & vbCrLf
            strXpertVal = strXpertVal & "3) Toxin B (+), Binary toxin (+), tcdC (+) : 고병원성 ribotype 027균주" & vbCrLf
            strXpertVal = strXpertVal & "   - Birnary toxin을 분비하면서 tcdC 유전자가 결손된 강한 독성을 보유한 C. difficile 균주"
        ElseIf strXpertCd = "L2759" Then
            strXpertVal = strXpertVal & vbCrLf
            strXpertVal = strXpertVal & "- 검사방법 : Real-Time PCR"
        ElseIf strXpertCd = "L2742" Then
            strXpertVal = Mid(strXpertVal, 1, Len(strXpertVal) - 2)
        ElseIf strXpertCd = "L2808" Then
            strXpertVal = strXpertVal & vbCrLf
            strXpertVal = strXpertVal & "- Test Method : Real-Time PCR" & vbCrLf
            strXpertVal = strXpertVal & vbCrLf
            strXpertVal = strXpertVal & "- Comment :" & vbCrLf
            strXpertVal = strXpertVal & "  PCR 검사는 검체 내 균수가 적거나 부적절한 검체 희석 또는 증폭 억제물질이 존재하는 경우" & vbCrLf
            strXpertVal = strXpertVal & "  위음성이 나올수 있습니다. 또한 PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이" & vbCrLf
            strXpertVal = strXpertVal & "  안되어 위양성의 가능성이 있습니다." & vbCrLf
            strXpertVal = strXpertVal & "  결과 해석 시 환자의 임상 양상과 연관지어 판단하시기 바랍니다."
        Else
        
        End If
        
        sParam = ""
        sParam = sParam & "<Table>"
        sParam = sParam & "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>"
        sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
        sParam = sParam & "<USERID><![CDATA[LIA]]></USERID>"
        sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
        sParam = sParam & "<TABLENAME><![CDATA[]]></TABLENAME>"
        sParam = sParam & "<P0><![CDATA[" & lsID & "]]></P0>"
        sParam = sParam & "<P1><![CDATA[" & strXpertCd & "]]></P1>"
        sParam = sParam & "<P2><![CDATA[" & strXpertVal & "]]></P2>"
        sParam = sParam & "<P3><![CDATA[]]></P3>"
        sParam = sParam & "<P4><![CDATA[" & gEquip & strEquip & "]]></P4>"
        sParam = sParam & "<P5><![CDATA[" & gIFUser & "]]></P5>"
        sParam = sParam & "<P6><![CDATA[]]></P6>"
        sParam = sParam & "<P7><![CDATA[]]></P7>"
        sParam = sParam & "<P8><![CDATA[]]></P8>"
        sParam = sParam & "<P9><![CDATA[]]></P9>"
        sParam = sParam & "</Table>"
    
    End If
    
    If sParam = "" Then Exit Function
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    Insert_Data = 1

    'Save_Raw_Data lsID & " : 서버 결과 전송 완료!"
    
    SQL = ""
    SQL = SQL & "Update pat_res set "
    SQL = SQL & "   sendflag    = 'C'               " & vbCrLf
    SQL = SQL & " where equipno = '" & gEquip & "'  " & vbCrLf
    SQL = SQL & "   and barcode = '" & lsID & "'    " & vbCrLf
    SQL = SQL & "   and examcode in (" & lsTransCode & ")"
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Exit Function

ErrHandle:
    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & SQL
    Resume Next
    
End Function

Sub Var_Clear()
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""
End Sub

Private Sub Picture1_Click()
    frmUser.Show 0
    
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    End If
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim lsRow As Integer
    
    If KeyCode = 13 Then
        lsRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = Trim(txtBarcode.Text) Then
                lsRow = i
                Exit For
            End If
        Next
        
        If lsRow = -1 Then
            lsRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < lsRow Then
                vasID.MaxRows = lsRow
            End If
        End If
        
        SetText vasID, Trim(txtBarcode.Text), lsRow, colBarCode
        
        If Trim(GetText(vasID, lsRow, colPID)) = "" Then
            Get_Sample_Info lsRow
        End If
        
        
    End If
End Sub

Private Sub txtEnd_GotFocus()
    SelectFocus txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtEnd) = False Then
            txtEnd.SetFocus
            Exit Sub
        End If
        cmdSend.SetFocus
    End If
End Sub

Private Sub txtHelp_Change()

End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
End Sub

Private Sub txtStart_GotFocus()
    SelectFocus txtStart
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtStart) = False Then
            txtStart.SetFocus
            Exit Sub
        End If
        txtEnd.SetFocus
    End If
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col = colRack Or Col = colPos Then
            vasSort vasID, colRack, colPos
        Else
            vasSort vasID, Col
        End If
    End If
    
    If Row < 0 Or Row > vasID.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasID.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsCnt As String
    Dim lsID As String
    Dim lsDate As String
    Dim lsTime As String
    Dim lsState As String
    Dim strExamDate As String
    
    
    Dim iRow As Long
    
    'cmdCall_Click
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
    
    strExamDate = Trim(GetText(vasID, Row, colDate))
    
    
'    If Trim(GetText(vasID, Row, colState)) = "결과" Then
'        lsState = "A"
'    ElseIf Trim(GetText(vasID, Row, colState)) = "완료" Then
'        lsState = "C"
'    End If
    'Local에서 불러오기
    ClearSpread vasRes
    
'    If Trim(GetText(vasID, Row, colPJumin)) = "F" Then
'        lsTime = Trim(GetText(vasID, Row, colPID))
'        If Len(lsTime) = 4 Then
'        Else
'            lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'        End If
'        SQL = "select a.equipcode, min(b.examcode), min(b.examname), a.result, b.seqno, a.resflag, a.result " & vbCrLf & _
'              " From qc_res a, equipexam b " & vbCrLf & _
'              "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'              "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'              "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
'              "  and a.levelname = '" & lsID & "' " & vbCrLf & _
'              "  and b.equipno = a.equipno " & vbCrLf & _
'              "  and b.equipcode = a.equipcode " & vbCrLf & _
'              "group by a.equipcode, a.result, b.seqno, a.resflag, a.result "
'        res = db_select_Vas(gLocal, SQL, vasRes)
'    End If
    

    '장비코드, 검사코드, 검사명, 결과, 순번
'    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
'          "from pat_res a, equipexam b " & vbCrLf & _
'          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
'          "  and a.examcode <> a.equipcode " & vbCrLf & _
'          "  and b.equipno = a.equipno " & vbCrLf & _
'          "  and b.equipcode = a.equipcode " & vbCrLf & _
'          "  and b.examcode = a.examcode"
'    res = db_select_Vas(gLocal, SQL, vasRes)
    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, max(b.seqno), a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "group by a.equipcode, a.examcode, b.examname, a.result,  a.refflag, a.result1 "

    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRes.DataRowCnt
        If Trim(GetText(vasRes, iRow, colRCheck)) <> "" Then
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 255, 0, 0
        Else
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 0, 0, 0
        End If
    Next iRow
    vasRes.MaxRows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarCode))
        
        If Trim(GetText(vasID, iRow, colPJumin)) = "F" Then
            If MsgBox("해당 QC 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
            
            lsTime = Trim(GetText(vasID, iRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
                
            Exit Sub
        End If
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        ClearSpread vasRes
    End If
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_DblClick colBarCode, lRow
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'Dim iRow As Long
'Dim lsID As String
'
'    If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'        Exit Sub
'    End If
'
'    iRow = Row
'
'    lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'    SQL = " Delete From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & lsID & "' "
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    DeleteRow vasID, iRow, iRow
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
    
    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String
    
    Dim sResult As String
    Dim sResult1 As String
    
    Dim i As Integer
    
    Dim sTotalVol As String
    
    Dim lsTime As String
    
    vasIDRow = vasID.ActiveRow
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    
    If KeyCode = vbKeyReturn Then

        If vasResCol = colResult Then
            
            If Trim(GetText(vasRes, vasResRow, colEquipCode)) = "88888" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
            
            ElseIf Trim(GetText(vasRes, vasResRow, colEquipCode)) = "99999" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
                
                If IsNumeric(sTotalVol) Then
                    lCCR = -1
                    sCCR = ""
                    sCrea_S = ""
                    sCrea_U = ""
                    sM_ALB_U = ""
                    sTP_U = ""
                    
                    i = 1
                    Do While i <= vasRes.DataRowCnt
                        Select Case Trim(GetText(vasRes, i, colExamCode))
                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                            
                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31094", "L31095" 'Protein 16hr, 8hr
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31111", "L31112", "L31123", "L3113" 'Creatinie 16hr, 8hr,24hr, BUN(24hr UR)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, "L31123", i, colExamCode
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100 / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L3041", "88888"   'Serum Creatinine
                            sCrea_S = Trim(GetText(vasRes, i, colResult1))
                            
                            'Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31121"   'CCR
                            sCCR = Trim(GetText(vasRes, i, colResult1))
                            lCCR = i
                        Case "L31171"   'Microalbumin(random)
                            sM_ALB_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31110"  'Creatinine(random)
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31090"   'Protein(random)
                            sTP_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31172"   'Microalbumin / creatinine (random urine)
                            lM_C_ratio = i
                        Case "L31172"   'protein / creatinie (random)
                            lP_C_ratio = i
                        End Select
                        i = i + 1
                    Loop
                    
                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCrea_U) = True And IsNumeric(sCrea_S) = True Then
                        sResult = Format(CCur(sCrea_U) * CCur(sTotalVol) / 1440 / CCur(sCrea_S), "0.000")
                        SetText vasRes, sResult, lCCR, colResult
                        SetText vasRes, sResult, lCCR, colResult1
                        Save_Local_One_1 vasIDRow, i, "A"
                    End If
                    
'                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
'                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "101", i, colEquipCode
'                            SetText vasRes, "L31172", i, colExamCode
'                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
'                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "102", i, colEquipCode
'                            SetText vasRes, "L31201", i, colExamCode
'                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
                End If
            Else
                
                If Trim(GetText(vasRes, vasIDRow, colPJumin)) = "F" Then
                
                    If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 수정 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                        Exit Sub
                    End If
                
                    lsTime = Trim(GetText(vasID, vasIDRow, colPID))
                    If Len(lsTime) = 4 Then
                    Else
                        lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
                    End If
                    
                    SQL = "update qc_res set result = '" & sResult & "' " & vbCrLf & _
                          "where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                          "  and examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                          "  and examtime = '" & lsTime & "' " & vbCrLf & _
                          "  and levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                          "  and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                    res = SendQuery(gLocal, SQL)
                
                    Exit Sub
                Else
                
                
                    sResult = Trim(GetText(vasRes, vasResRow, colResult))
                    If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
                        sResult = Trim(GetText(vasRes, vasResRow, colResult))
                        
                        SQL = " update pat_res set " & vbCrLf & _
                              " Result = '" & sResult & "' " & vbCrLf & _
                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                              " And equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                        If res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
        
                        'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
        
                    End If
                End If
            End If
            
            
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Trim(GetText(vasID, vasIDRow, colPJumin)) = "F" Then
        
            If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
        
            lsTime = Trim(GetText(vasID, vasIDRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                  " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
            res = SendQuery(gLocal, SQL)
        
            Exit Sub
        End If
        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' " & vbCrLf & _
              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        DeleteRow vasRes, vasResRow, vasResRow
    
    End If
End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function


Private Sub wSck_Close()
        
    wSck.Close
    wSck.LocalPort = CInt(5150)
    wSck.Listen

    lblStatus.Caption = "TCP " & "5150" & " 포트에 연결 되었습니다"

End Sub

Private Sub wSck_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        lblStatus.Caption = "TCP " & "5150" & " 포트로 연결 되었습니다"
    End If
            
End Sub

Private Sub wSck_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    
    wSck.GetData strText
    Save_Raw_Data "[Rx]" & strText
    
    '-- 컴파일시 제외할 것!!
    'strText = Replace(strText, vbLf, "")
    
    Call RcvSocketData(strText)

End Sub

