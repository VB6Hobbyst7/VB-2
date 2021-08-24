VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   " URIT3000 Interface Program"
   ClientHeight    =   10530
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   14925
   FillColor       =   &H0000FFFF&
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "Hidden"
      Height          =   10095
      Left            =   15390
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   17715
      Begin VB.TextBox txtToday 
         Appearance      =   0  '평면
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   66
         Text            =   "2002/02/18"
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txtMsg 
         ForeColor       =   &H000000C0&
         Height          =   825
         Left            =   5850
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   57
         Top             =   6780
         Visible         =   0   'False
         Width           =   11745
      End
      Begin VB.TextBox txtDate 
         Height          =   405
         Left            =   11310
         TabIndex        =   56
         Top             =   5580
         Width           =   2325
      End
      Begin VB.TextBox txtAll 
         Height          =   375
         Left            =   7920
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   6240
         Width           =   2055
      End
      Begin VB.TextBox txtTemp 
         Height          =   375
         Left            =   7920
         TabIndex        =   54
         Top             =   5130
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "취소"
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
         Left            =   6180
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   7740
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtBuff 
         Height          =   1350
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   6840
         Visible         =   0   'False
         Width           =   5010
      End
      Begin VB.TextBox txtBuff2 
         Height          =   1305
         Left            =   240
         TabIndex        =   43
         Top             =   8340
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   7
         Top             =   600
         Width           =   1515
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3735
         Left            =   150
         TabIndex        =   10
         Top             =   1050
         Visible         =   0   'False
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   6588
         _Version        =   131072
         BackColor       =   16761024
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2055
            TabIndex        =   37
            Text            =   "123"
            Top             =   1185
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   2055
            TabIndex        =   26
            Text            =   "123"
            Top             =   1695
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   2055
            TabIndex        =   25
            Text            =   "123"
            Top             =   2205
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   2055
            TabIndex        =   24
            Text            =   "123"
            Top             =   2715
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   2055
            TabIndex        =   23
            Text            =   "123"
            Top             =   3195
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   5
            Left            =   5400
            TabIndex        =   22
            Text            =   "123"
            Top             =   1185
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   6
            Left            =   5400
            TabIndex        =   21
            Text            =   "123"
            Top             =   1695
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   7
            Left            =   5400
            TabIndex        =   20
            Text            =   "123"
            Top             =   2205
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   8
            Left            =   5400
            TabIndex        =   19
            Text            =   "123"
            Top             =   2715
            Width           =   1105
         End
         Begin VB.TextBox txtMain 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   9
            Left            =   5400
            TabIndex        =   18
            Text            =   "123"
            Top             =   3225
            Width           =   1105
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   1065
            Left            =   75
            TabIndex        =   11
            Top             =   60
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   1879
            _Version        =   131072
            BackColor       =   12632319
            BorderWidth     =   1
            BevelInner      =   1
            Begin VB.TextBox txtMain1 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   1980
               TabIndex        =   15
               Text            =   "123"
               Top             =   90
               Width           =   1215
            End
            Begin VB.TextBox txtMain1 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   1980
               TabIndex        =   14
               Text            =   "Positive"
               Top             =   570
               Width           =   1215
            End
            Begin VB.TextBox txtMain1 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   2
               Left            =   5175
               TabIndex        =   12
               Text            =   "Negative"
               Top             =   570
               Width           =   1215
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   420
               Left            =   3315
               TabIndex        =   13
               Top             =   570
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   741
               _Version        =   131072
               Caption         =   "HBs Ag(RPHA)"
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   420
               Left            =   105
               TabIndex        =   16
               Top             =   570
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   741
               _Version        =   131072
               Caption         =   "HBs Ab(RPHA)"
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   420
               Left            =   105
               TabIndex        =   17
               Top             =   90
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   741
               _Version        =   131072
               Caption         =   "HbA1c"
               BorderWidth     =   0
               BevelOuter      =   1
               BevelInner      =   2
            End
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   420
            Left            =   3405
            TabIndex        =   27
            Top             =   2715
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Specific Gravity"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   420
            Left            =   3405
            TabIndex        =   28
            Top             =   2205
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Glucose"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   420
            Left            =   3405
            TabIndex        =   29
            Top             =   1695
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Leukocytes"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   420
            Left            =   3405
            TabIndex        =   30
            Top             =   1185
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Nitrite"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   420
            Left            =   3405
            TabIndex        =   31
            Top             =   3225
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "pH"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   420
            Left            =   75
            TabIndex        =   32
            Top             =   3225
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Protein"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   420
            Left            =   75
            TabIndex        =   33
            Top             =   2715
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "RBC"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   420
            Left            =   75
            TabIndex        =   34
            Top             =   2205
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Ketone"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   420
            Left            =   75
            TabIndex        =   35
            Top             =   1695
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Blllrubin"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   420
            Index           =   0
            Left            =   60
            TabIndex        =   36
            Top             =   1185
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   741
            _Version        =   131072
            Caption         =   "Urobllinogen"
            BorderWidth     =   0
            BevelOuter      =   1
            BevelInner      =   2
         End
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1095
         Left            =   7080
         TabIndex        =   38
         Top             =   420
         Visible         =   0   'False
         Width           =   2865
         _Version        =   393216
         _ExtentX        =   5054
         _ExtentY        =   1931
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
         SpreadDesigner  =   "frmInterface.frx":0442
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   960
         Left            =   7020
         TabIndex        =   39
         Top             =   1620
         Visible         =   0   'False
         Width           =   1605
         _Version        =   393216
         _ExtentX        =   2831
         _ExtentY        =   1693
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
         SpreadDesigner  =   "frmInterface.frx":0663
      End
      Begin FPSpread.vaSpread vasOrderTemp 
         Height          =   990
         Left            =   6990
         TabIndex        =   40
         Top             =   2820
         Visible         =   0   'False
         Width           =   1905
         _Version        =   393216
         _ExtentX        =   3360
         _ExtentY        =   1746
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
         SpreadDesigner  =   "frmInterface.frx":41B4
      End
      Begin FPSpread.vaSpread vasOrderBuf 
         Height          =   750
         Left            =   7080
         TabIndex        =   41
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   393216
         _ExtentX        =   2355
         _ExtentY        =   1323
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
         SpreadDesigner  =   "frmInterface.frx":7CEB
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1695
         Left            =   270
         TabIndex        =   42
         Top             =   4920
         Visible         =   0   'False
         Width           =   5895
         _Version        =   393216
         _ExtentX        =   10398
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
         SpreadDesigner  =   "frmInterface.frx":B822
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   5700
         Top             =   8370
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   3
         DTREnable       =   -1  'True
         InBufferSize    =   4096
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         DataBits        =   7
         StopBits        =   2
         EOFEnable       =   -1  'True
      End
      Begin MSCommLib.MSComm MSComm2 
         Left            =   6420
         Top             =   8370
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InBufferSize    =   4096
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   4455
         Left            =   10200
         TabIndex        =   53
         Top             =   510
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":BA43
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
         Height          =   195
         Left            =   4860
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "현재 검사수"
         Height          =   195
         Left            =   3405
         TabIndex        =   47
         Top             =   645
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblToday 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "100,000"
         Height          =   195
         Left            =   4230
         TabIndex        =   46
         Top             =   270
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오늘 검사수"
         Height          =   195
         Left            =   2970
         TabIndex        =   45
         Top             =   270
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   8
         Top             =   645
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9600
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Width           =   14790
      Begin VB.CheckBox ChkAll 
         Height          =   255
         Left            =   780
         TabIndex        =   58
         Top             =   810
         Width           =   165
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 불러오기"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         Style           =   1  '그래픽
         TabIndex        =   49
         Top             =   180
         Width           =   2505
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "추가"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12060
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "결과받기"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5520
         TabIndex        =   4
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdCall 
         Caption         =   "불러오기"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4020
         TabIndex        =   3
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrt 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10560
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   1485
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8745
         Left            =   270
         TabIndex        =   51
         Top             =   720
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   15425
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         EditEnterAction =   8
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   16777215
         MaxCols         =   50
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmInterface.frx":FF4A
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8715
         Left            =   7410
         TabIndex        =   52
         Top             =   750
         Width           =   6945
         _Version        =   393216
         _ExtentX        =   12250
         _ExtentY        =   15372
         _StockProps     =   64
         ColHeaderDisplay=   1
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   16777215
         MaxCols         =   13
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmInterface.frx":124B3
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   1296
      _Version        =   131072
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "   URIT3000 INTERFACE"
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.FileListBox FileURIT 
         Height          =   675
         Left            =   0
         Pattern         =   "*.txt"
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7620
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11220
         TabIndex        =   63
         Top             =   150
         Width           =   1125
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13500
         TabIndex        =   62
         Top             =   150
         Width           =   1125
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12360
         TabIndex        =   61
         Top             =   150
         Width           =   1125
      End
      Begin VB.CommandButton cmd_Trans 
         Caption         =   "선택전송"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10050
         TabIndex        =   60
         Top             =   150
         Width           =   1155
      End
      Begin Threed.SSPanel sspMode 
         Height          =   525
         Left            =   8880
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   926
         _Version        =   131072
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "전송모드"
         BevelWidth      =   3
         BorderWidth     =   5
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   4560
         TabIndex        =   67
         Top             =   210
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63438848
         CurrentDate     =   40457
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   68
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pp"
      Visible         =   0   'False
      Begin VB.Menu subDel 
         Caption         =   "삭제"
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
Const colRack = 2
Const ColPos = 3
Const colSampleNo = 4
Const colPID = 5
Const colPName = 6
Const colJumin = 7
Const colPSex = 8
Const colPAge = 9
Const colOCnt = 10
'Const colRCnt = 11
Const colState = 11
Const colReceNo = 12    '2004/07/15 이상은 추가
Const colReqDate = 13   '접수일자



Const colEquipExam = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13

Const TXT_URO = 0
Const TXT_BIL = 1
Const TXT_KET = 2
Const TXT_RBC = 3
Const TXT_PRO = 4
Const TXT_NIT = 5
Const TXT_LEU = 6
Const TXT_GLU = 7
Const TXT_SPE = 8
Const TXT_PH1 = 9

Const TXT_HBA = 0
Const TXT_HBG = 1
Const TXT_HBB = 2

Dim ConfirmData As String

Public gPID As String
Public gTestID As String
Public gSpecID As String
Public glRow As Long
Public gCount As String
Public gOCnt As Integer
Public gOCnt_1 As Integer
Public gRCnt As Integer
Public gCheck As String

'===============================
Dim strBufferData As String

Private Type typeXMLInData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLInData As typeXMLInData

Private Type typeXMLOutData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLOutData As typeXMLOutData

Dim blnSameRecord As Boolean

'===============================
Const SPCLEN As Integer = 10

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
Const GS  As String = ""


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================

Dim gRow As Long


Private Sub chkAll_Click()
    Dim iRow As Integer
    
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

Private Sub cmd_Trans_Click()
''''선택전송
'''Dim vasIDRow As Integer
'''Dim vasResRow As Integer
'''Dim iRow As Integer
'''Dim liRet As Integer
'''Dim iNumber As Integer
'''
'''    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
'''        Exit Sub
'''    End If
'''
'''    If txtUID.Text = "" Then
'''        MsgBox "사용자 확인을 해 주십시오"
'''        txtUID.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If (vasID.DataRowCnt < 1) Or (vasRes.DataRowCnt < 1) Then
'''        MsgBox "저장할 데이터가 없습니다."
'''        Exit Sub
'''    End If
'''
'''    'db_BeginTran gServer
'''
'''    For vasIDRow = 1 To vasID.DataRowCnt
'''        vasID.Col = 1
'''        vasID.Row = vasIDRow
'''        If vasID.Value = 1 Then
'''
'''            liRet = -1
'''            If Trim(GetText(vasID, vasIDRow, colPID)) <> "" Then
'''                liRet = Insert_Data(vasIDRow)
'''            End If
'''
'''            If liRet = 1 Then
'''                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
'''                SetText vasID, "완료", vasIDRow, colState
'''            Else
'''                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
'''                SetText vasID, "실패", vasIDRow, colState
'''            End If
'''        Else
'''
'''        End If
'''    Next vasIDRow
'''
    
    '선택전송
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim FindFile As String
    
    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If

'''    If txtUID.Text = "" Then
'''        MsgBox "사용자 확인을 해 주십시오"
'''        txtUID.SetFocus
'''        Exit Sub
'''    End If
    
    If (vasID.DataRowCnt < 1) Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
    
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        If vasID.Value = 1 Then
            liRet = -1
            If Trim(GetText(vasID, vasIDRow, colPID)) <> "" Then
               liRet = Make_XML_File(vasIDRow)
            End If
            
            If liRet = 1 Then
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "전송완료", vasIDRow, colState
                
                FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                If FindFile <> "" Then
                    Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '전송완료가 됐을때 파일지우기
                End If
                SQL = "update pat_res set sendflag = 'B' where barcode = '" & GetText(vasID, vasIDRow, 5) & "' and examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "'"
                res = SendQuery(gLocal, SQL)
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", vasIDRow, colState
            End If
        Else
        
        End If
    Next vasIDRow
    
    If XmlTxtHead = "" Then
    XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                 "<?xml-stylesheet type=""text/xsl"" href=""C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare검사정보>"
    End If
    If XmlTxtTail = "" Then
    XmlTxtTail = "</UBCare검사정보>"
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    SaveXMLFile XMLAllTxt
    
End Sub

Public Sub SaveXMLFile(argSQL As String, Optional argFlag As Integer = 0)
'argSQL의 내용을 파일로 저장
    Dim FilNum, FilNum1
    Dim FindFile As String
    Dim TxtString1 As String
    Dim AllString1 As String
    Dim i As Long
    
    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out.xml")
    
    
    If FindFile <> "" Then
'        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
        FilNum1 = FreeFile
        Open "C:\UBCare\SINAI\IF\ExamIF_out.xml" For Input As FilNum1
        Do While Not EOF(FilNum1)
            Input #FilNum1, TxtString1
            AllString1 = AllString1 & TxtString1
        Loop
        Close #FilNum1
        i = InStr(1, AllString1, "</UBCare검사정보>")
        XmlBody = Mid(AllString1, 1, i - 1)
        argSQL = XmlBody & argSQL & XmlTxtTail
        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
    Else
        argSQL = XmlTxtHead & argSQL & XmlTxtTail
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    
    FilNum = FreeFile
    
    
    If argFlag = 0 Then
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Output As FilNum
    Else
        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
    argSQL = ""
End Sub

Function Make_XML_File(asRow) As Integer
    Dim FilNum
    Dim FilNum2
    Dim TxtString As String
    Dim ResultString As String
    Dim TxtRece As String
    Dim i As Long
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim pResult As String
    Dim pExamdate As String
    Dim pOpinion As String
    Dim TxtPat As String
    Dim IOGubun As String
    Dim TestNum As String
    Make_XML_File = -1

'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If

'    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    ClearSpread vasResTemp
    SQL = "select  pid, examcode, recedate, barcode,pname, pjumin, examdate, '', seqno,result " & vbCrLf & _
          "from pat_res where examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & CR & _
          " and result <> '' " & CR & _
          " And equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasID, asRow, 5)) & "'"
    SQL = SQL & "order by seqno "
    res = db_select_Vas(gLocal, SQL, vasResTemp)

'          " and subcode <> '' " & CR & _

    For i = 1 To vasResTemp.DataRowCnt
'    XMLAllTxt = ""
        PID = Trim(GetText(vasResTemp, i, 1))
        PExamCode = Trim(GetText(vasResTemp, i, 2))
        PReceDate = Trim(GetText(vasResTemp, i, 3))
        PChartNum = Trim(GetText(vasResTemp, i, 4))
        PName = Trim(GetText(vasResTemp, i, 5))
        PJumin = Mid(Trim(GetText(vasResTemp, i, 6)), 1, 6) & "-" & Mid(Trim(GetText(vasResTemp, i, 6)), 7)
        pExamdate = Trim(GetText(vasResTemp, i, 7))
        IOGubun = Trim(GetText(vasResTemp, i, 8))
        TestNum = Trim(GetText(vasResTemp, i, 9))
        pResult = Trim(GetText(vasResTemp, i, 10))
        XMLAllTxt = XMLAllTxt & "<검사><업체>ACK</업체><요양기관번호>38341948</요양기관번호><차트번호>" & PChartNum & "</차트번호><수진자명>" & PName & "</수진자명><주민등록번호>" & PJumin & "</주민등록번호><내원번호>" & PID & "</내원번호><의뢰일>" & PReceDate & "</의뢰일><검사번호>" & TestNum & "</검사번호><검사ID>" & PExamCode & "</검사ID><업체검사ID></업체검사ID><검체></검체><결과치>" & pResult & "</결과치><참조치></참조치><소견></소견><결과일>" & pExamdate & "</결과일><입원외래구분>" & IOGubun & "</입원외래구분></검사>"
    Next
    

    Make_XML_File = 1
End Function

Function Insert_Data(argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim i As Integer
    
    Dim sNumber As String
    
    
    Insert_Data = -1
       
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, result, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where examdate = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' "
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    
    '서버로 결과값 저장하기
    For i = 1 To vasResTemp.DataRowCnt
        
        SQL = " Select OdrOcmNum" & CR & _
              " From OdrInf " & CR & _
              " Where OdrDtm  like '" & Format(txtToday.Text, "yyyymmdd") & "%' " & CR & _
              " And odrchtnum = " & Trim(GetText(vasID, argSpcRow, colPID)) & "  " & CR & _
              " And odrcod =  '" & Trim(GetText(vasResTemp, i, 2)) & "' "
        res = db_select_Col(gServer, SQL)
        
        If gReadBuf(0) <> "" Then
            sNumber = Trim(gReadBuf(0))
            
            SQL = " Update ResInf Set " & CR & _
                  " ResRltVal = '" & Trim(GetText(vasResTemp, i, 3)) & "', " & CR & _
                  " ResUpdDtm = '" & Format(GetDateFull, "yyyymmddhhmm") & "', " & CR & _
                  " ResUpdUid = '" & Trim(txtUID.Text) & "' " & CR & _
                  " Where ResOcmNum = " & sNumber & "" & CR & _
                  " And ResLabCod = '" & Trim(GetText(vasResTemp, i, 2)) & "' "
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                Exit Function
            End If
            
        End If
    Next i

    Insert_Data = 1
    
End Function

Private Sub cmdAdd_Click()
    
    Dim i As Integer
    
    For i = 0 To 9
        txtMain(i).Text = ""
    Next
    
    For i = 0 To 2
        txtMain1(i).Text = ""
    Next
    
    SSPanel2.Visible = True
    txtMain1(TXT_HBA).SetFocus

End Sub


Private Sub cmdCall_Click()
    ClearSpread vasID
    
    SQL = " Select barcode, pname, '', psex, page " & CR & _
          " From pat_res " & CR & _
          " Where examdate = '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
          " Group By barcode, pname, psex, page "
    res = db_select_Vas(gLocal, SQL, vasID, , 5)
    
    
        
End Sub

Private Sub cmdCancel_Click()
    vasOrder.Row = 1
    vasOrder.Row2 = vasOrder.MaxRows
    vasOrder.Col = 1
    vasOrder.Col2 = vasOrder.MaxCols
    vasOrder.BlockMode = True
    vasOrder.Action = 3
    vasOrder.BlockMode = False
    
    vasOrderTemp.Row = 1
    vasOrderTemp.Row2 = vasOrderTemp.MaxRows
    vasOrderTemp.Col = 1
    vasOrderTemp.Col2 = vasOrderTemp.MaxCols
    vasOrderTemp.BlockMode = True
    vasOrderTemp.Action = 3
    vasOrderTemp.BlockMode = False
    
    vasOrderBuf.Row = 1
    vasOrderBuf.Row2 = vasOrderTemp.MaxRows
    vasOrderBuf.Col = 1
    vasOrderBuf.Col2 = vasOrderTemp.MaxCols
    vasOrderBuf.BlockMode = True
    vasOrderBuf.Action = 3
    vasOrderBuf.BlockMode = False
    
End Sub

Private Sub cmdClear_Click()
Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasID
'    ClearSpread vasWork
'    ClearSpread vasList

'    vasID.MaxRows = 1
        
    vasActiveCell vasID, 1, colPID

    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 1
    vasRes.OperationMode = 0
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
    Exit Sub
End Sub

Private Sub cmdConfig_Click()
'    frmConfig.SSPanel_machine.Caption = "BT2000 plus"
    frmConfig.SSPanel_machine.Caption = "URIT8021A"
    frmConfig.Show 1
End Sub

Private Sub cmdPrt_Click()
'    Call Print_Report
'    Call Print_Report
End Sub

Private Sub cmdReceive_Click()

    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strBuffer   As String
    Dim i           As Long
    Dim varBuffer   As Variant

    FileURIT.Path = gMachPath
    FileURIT.Refresh

    For intIdx = 0 To FileURIT.ListCount - 1
        
        FileURIT.ListIndex = intIdx
        
        '===== 조회기간에 맞는것만 1 =================================================================
        If FileURIT.FileName = "NameResult.txt" Then
            strSrcfile = FileURIT.Path & "\" & FileURIT.FileName   ' 원본 파일 이름을 정의합니다.
            
            Open strSrcfile For Input As #9
        
            strBuffer = ""
        
            Do While Not EOF(9)
                strBuffer = strBuffer & Input(1, #9)
            Loop
        
            Close #9
            
            If InStr(strBuffer, vbCrLf) = 0 Then
                strBuffer = strBuffer & vbCrLf
            End If
            
            varBuffer = Split(strBuffer, vbCrLf)
            
            ReDim Preserve strRecvData(UBound(varBuffer))
            
            For i = 0 To UBound(varBuffer) '- 1
                'Debug.Print varBuffer(i)
                If Trim(Mid(varBuffer(i), 13, 12)) <> "" Then
                    strRecvData(i) = varBuffer(i)
                    
                Else
                    '-- test
                    'strRecvData(i) = varBuffer(i)
                End If
            Next i
            
            
'            Call EditRcvDataURIT
            Call URIT8021A(strRecvData)
            
            strDestFile = App.Path & "\Log\" & Format(Now, "yyyymmddhhmm") & ".txt"
            '원본을 대상에 복사
            FileCopy strSrcfile, strDestFile
            
            Kill strSrcfile
            
        End If
    Next
    
End Sub


Private Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = "C:\UBCare\SINAI\IF\ExamIF_In.xml"

'    Open strPath For Input As #3
'
'    strBuffer = ""
'    Do While Not EOF(3)
'        strBuffer = strBuffer & Input(1, #3)
'    Loop
'
'    Close #3
'
    
    '1라인씩 가져오기 MSDN내용
    Dim TextLine
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1) ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        'Debug.Print TextLine ' 직접 실행 창에 출력합니다.
        strBuffer = strBuffer & TextLine
'        txtBuffer = txtBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    '파일전체 읽기 MSDN내용
    'Dim FileLength
    'Open "TESTFILE" For Input As #1   ' 파일을 엽니다.
    'FileLength = LOF(1)   ' 파일의 길이를 구합니다.
    'Close #1   ' 파일을 닫습니다.
    
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    
    Next
    
'    f_subSet_XMLWorkList = Split(strTmp, "</검사>")
    strTmp = Replace(strTmp, "<검사>", ""): strTmp = Replace(strTmp, "</검사>", "|")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<요양기관번호>", ""): strTmp = Replace(strTmp, "</요양기관번호>", ",")
    strTmp = Replace(strTmp, "<차트번호>", ""): strTmp = Replace(strTmp, "</차트번호>", ",")
    strTmp = Replace(strTmp, "<수진자명>", ""): strTmp = Replace(strTmp, "</수진자명>", ",")
    strTmp = Replace(strTmp, "<주민등록번호>", ""): strTmp = Replace(strTmp, "</주민등록번호>", ",")
    strTmp = Replace(strTmp, "<내원번호>", ""): strTmp = Replace(strTmp, "</내원번호>", ",")
    strTmp = Replace(strTmp, "<의뢰일>", ""): strTmp = Replace(strTmp, "</의뢰일>", ",")
    strTmp = Replace(strTmp, "<검사번호>", ""): strTmp = Replace(strTmp, "</검사번호>", ",")
    strTmp = Replace(strTmp, "<검사ID>", ""): strTmp = Replace(strTmp, "</검사ID>", ",")
    strTmp = Replace(strTmp, "<업체검사ID>", ""): strTmp = Replace(strTmp, "</업체검사ID>", ",")
    strTmp = Replace(strTmp, "<검체>", ""): strTmp = Replace(strTmp, "</검체>", ",")
    strTmp = Replace(strTmp, "<결과치>", ""): strTmp = Replace(strTmp, "</결과치>", ",")
    strTmp = Replace(strTmp, "<참조치>", ""): strTmp = Replace(strTmp, "</참조치>", ",")
    strTmp = Replace(strTmp, "<소견>", ""): strTmp = Replace(strTmp, "</소견>", ",")
    strTmp = Replace(strTmp, "<결과일>", ""): strTmp = Replace(strTmp, "</결과일>", ",")
    strTmp = Replace(strTmp, "<업체>", ""): strTmp = Replace(strTmp, "</업체>", ",")
    strTmp = Replace(strTmp, "<입원외래구분>", ""): strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    blnSameRecord = True
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function


Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

'Private Sub cmdRegi_Click()
'
'    Dim varTmp  As Variant
'    Dim intRow1 As Integer, intRow2 As Integer
'    Dim intIdx  As Integer
'    Dim Rev     As Long
'    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
'    Dim blnFlag As Boolean
'    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
'    Dim strWDate As String
'    Dim strEqpCd    As String
'    Dim tmpDate     As String
'    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD, strSQNO As String
'    Dim strData(20) As String
'
'    blnFlag = False
'
'    With vasWork
'        For intRow1 = 1 To .MaxRows
'            .GetText 1, intRow1, varTmp
'            If Trim$(varTmp) = "1" Then
'                .GetText 2, intRow1, varTmp:    strData(2) = Trim$(varTmp)
'                .GetText 3, intRow1, varTmp:    strData(3) = Trim$(varTmp)
'                .GetText 4, intRow1, varTmp:    strData(4) = Trim$(varTmp)
'                .GetText 5, intRow1, varTmp:    strData(5) = Trim$(varTmp)
'                .GetText 6, intRow1, varTmp:    strData(6) = Trim$(varTmp)
'                .GetText 7, intRow1, varTmp:    strData(7) = Trim$(varTmp)
'                .GetText 8, intRow1, varTmp:    strData(8) = Trim$(varTmp)
'                .GetText 9, intRow1, varTmp:    strData(9) = Trim$(varTmp)
'                .GetText 10, intRow1, varTmp:   strData(10) = Trim$(varTmp)
'                .GetText 11, intRow1, varTmp:   strData(11) = Trim$(varTmp)
'                .GetText 12, intRow1, varTmp:   strData(12) = Trim$(varTmp)
'                .GetText 13, intRow1, varTmp:   strData(13) = Trim$(varTmp)
'                .GetText 14, intRow1, varTmp:   strData(14) = Trim$(varTmp)
'                .GetText 15, intRow1, varTmp:   strData(15) = Trim$(varTmp)
'                .GetText 16, intRow1, varTmp:   strData(16) = Trim$(varTmp)
'                .GetText 17, intRow1, varTmp:   strData(17) = Trim$(varTmp)
'                .GetText 18, intRow1, varTmp:   strData(18) = Trim$(varTmp)
'                .GetText 19, intRow1, varTmp:   strData(19) = Trim$(varTmp)
'
'                .Row = intRow1: .Col = 1: .ForeColor = vbRed
'                                .Col = 2: .ForeColor = vbRed
'                                .Col = 3: .ForeColor = vbRed
'                                .Col = 4: .ForeColor = vbRed
'                                .Col = 5: .ForeColor = vbRed
'                                .Col = 6: .ForeColor = vbRed
'                                .Col = 7: .ForeColor = vbRed
'                                .Col = 8: .ForeColor = vbRed
'                                .Col = 9: .ForeColor = vbRed
'                                .Col = 10: .ForeColor = vbRed
'
'                intRow2 = SeqSearch(vasList, strData(4), 4)
'                If intRow2 < 1 Then
'                    intRow2 = SeqNullSearch(vasList, strData(4), 4)
'                    If intRow2 < 1 Then
'                        vasList.MaxRows = vasList.MaxRows + 1
'                        vasList.RowHeight(vasList.MaxRows) = 12
'                        intRow2 = vasList.MaxRows
'                    End If
'
'                    'blnFlag = False
'
'                    'If blnFlag = True Then
'                        vasList.SetText 2, intRow2, strData(2)
'                        vasList.SetText 3, intRow2, strData(3)
'                        vasList.SetText 4, intRow2, strData(4)
'                        vasList.SetText 5, intRow2, strData(5)
'                        vasList.SetText 6, intRow2, strData(6)
'                        vasList.SetText 7, intRow2, strData(7)
'                        vasList.SetText 8, intRow2, strData(8)
'                        vasList.SetText 9, intRow2, strData(9)
'                        vasList.SetText 10, intRow2, strData(10)
'                        vasList.SetText 11, intRow2, strData(11)
'                        vasList.SetText 12, intRow2, strData(12)
'                        vasList.SetText 13, intRow2, strData(13)
'                        vasList.SetText 14, intRow2, strData(14)
'                        vasList.SetText 15, intRow2, strData(15)
'                        vasList.SetText 16, intRow2, strData(16)
'                        vasList.SetText 17, intRow2, strData(17)
'                        vasList.SetText 18, intRow2, strData(18)
'                        vasList.SetText 19, intRow2, strData(19)
'
'                        vasList.Row = intRow2:
'                        vasList.Col = 2:
'                        vasList.ForeColor = vbRed
'                    'Else
'                    '    vasList.MaxRows = vasList.MaxRows - 1
'                    'End If
'                End If
'
'                .SetText 1, intRow1, ""
'            End If
'        Next
'    End With
'
'End Sub

'Private Sub cmdSearch_Click()
'    Dim sSch1, sSch2 As String
'    Dim iRow As Integer
'    Dim i, X As Long
'    Dim sCnt As String
'    Dim sExamCode As String
'    Dim sExamName As String
'    Dim FilNum
'    Dim TxtString As String
'    Dim TxtRece As String
'    Dim PChartNum As String
'    Dim PName As String
'    Dim PJumin As String
'    Dim PID As String
'    Dim PExamCode As String
'    Dim PReceDate As String
'    Dim PAge As String
'    Dim pSex As String
'    Dim STxt, NumTxt As Long
'    Dim SQL As String
'    Dim PEquipno As String
'    Dim PExamname As String
'    Dim PEquipCode As String
'    Dim j As Long
'    Dim BarFlag As Integer
'    Dim TxtPat As String
'    Dim TestNum, IOGubun As String
'    Dim FindFile As String
'    Dim StartDate As String
'    Dim EndDate As String
'    Dim varXML      As Variant
'    Dim varTmp      As Variant
'    Dim strBarno As String
'    Dim intCnt As Integer
'    Dim pGrid_Point As Integer
'    Dim sList As Integer
'
'    ClearSpread vasWork
'
'
'    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
''    Exit Sub
'
'    If blnSameRecord = False Then
'        MsgBox Format(dtpStartDt.Value, "yyyy-mm-dd") & "일 에서  " & Format(dtpStopDt.Value, "yyyy-mm-dd") & "일까지의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
'        Exit Sub
'    End If
'
'    If UBound(varXML) <= 1 Then
'        MsgBox dtpStartDt.Value & "일 에서  " & dtpStopDt.Value & "일까지의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
'        Exit Sub
'    Else
'        strBarno = ""
'
'        With vasWork
'            '.Visible = False
'            For intCnt = 0 To UBound(varXML) - 1
'                varTmp = Split(varXML(intCnt), ",")
'
'                XMLInData.Company = varTmp(0)
'                XMLInData.HospCode = varTmp(1)
'                XMLInData.ChartNo = varTmp(2)
'                XMLInData.PatName = varTmp(3)
'                XMLInData.PatJumin = varTmp(4)
'                XMLInData.PatNo = varTmp(5)
'                XMLInData.CommDate = varTmp(6)
'                XMLInData.ExamNo = varTmp(7)
'                XMLInData.ExamID = varTmp(8)
'                XMLInData.ComExamID = varTmp(9)
'                XMLInData.Specimen = varTmp(10)
'                XMLInData.Result = varTmp(11)
'                XMLInData.Reference = varTmp(12)
'                XMLInData.Remark = varTmp(13)
'                XMLInData.RsltDate = varTmp(14)
'                XMLInData.IOFlag = varTmp(15)
'
'                'If strBarno <> XMLInData.ChartNo Then
'                If XMLInData.CommDate >= Format(dtpStartDt.Value, "yyyymmdd") And XMLInData.CommDate <= Format(dtpStopDt.Value, "yyyymmdd") Then
'                    pGrid_Point = SeqSearch(vasList, XMLInData.ExamNo, 4)
'
'                    If pGrid_Point = 0 Then
'                        pGrid_Point = SeqNullSearch(vasWork, XMLInData.ExamNo, 4)
'                        If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
'                        .RowHeight(-1) = 12
'                    End If
'
'                    '-- 속도향상을 위해 쿼리문 지우기
'                    SQL = "select equipno, equipcode, examname from equipexam where examcode = '" & XMLInData.ExamID & "' "
'                    res = db_select_Col(gLocal, SQL)
'
'                    If res > 0 Then
'
'                        PEquipno = gReadBuf(0)
'                        PEquipCode = gReadBuf(1)
'                        PExamname = gReadBuf(2)
'
'                        'PEquipno = "BT2000"
'                        'PEquipCode = "Anti-ccp"
'                        'PExamname = "항CCP항체"
'
'                        .SetText 1, pGrid_Point, "0"
'                        .SetText 2, pGrid_Point, XMLInData.Company
'                        .SetText 3, pGrid_Point, XMLInData.HospCode
'                        .SetText 4, pGrid_Point, XMLInData.ChartNo
'                        .SetText 5, pGrid_Point, XMLInData.PatName
'                                    PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
'                                    Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
'                        .SetText 6, pGrid_Point, gPatGen.Sex
'                        .SetText 7, pGrid_Point, gPatGen.Age
'                        .SetText 8, pGrid_Point, XMLInData.PatJumin
'                        .SetText 9, pGrid_Point, XMLInData.PatNo
'                        .SetText 10, pGrid_Point, XMLInData.CommDate
'                        .SetText 11, pGrid_Point, XMLInData.ExamNo
'                        .SetText 12, pGrid_Point, XMLInData.ExamID
'                        .SetText 13, pGrid_Point, XMLInData.ComExamID
'                        .SetText 14, pGrid_Point, XMLInData.Specimen
'                        .SetText 15, pGrid_Point, XMLInData.Result
'                        .SetText 16, pGrid_Point, XMLInData.Reference
'                        .SetText 17, pGrid_Point, XMLInData.Remark
'                        .SetText 18, pGrid_Point, XMLInData.RsltDate
'                        .SetText 19, pGrid_Point, XMLInData.IOFlag
'
'
'                        '-- 오른쪽리스트에 있으면 붉은색
'                        With vasList
'                            For sList = 1 To vasList.DataRowCnt
'                                vasList.Row = sList
'                                vasList.Col = 4
'                                If Trim(vasList.Text) = Trim(XMLInData.ChartNo) Then
'                                    vasWork.Row = pGrid_Point: vasWork.Col = 1: vasWork.ForeColor = vbRed: vasWork.Value = "0"
'                                                               vasWork.Col = 2: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 3: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 4: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 5: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 6: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 7: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 8: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 9: vasWork.ForeColor = vbRed
'                                                               vasWork.Col = 10: vasWork.ForeColor = vbRed
'                                    Exit For
'                                Else
'                                    vasWork.Row = pGrid_Point
'                                    vasWork.Col = 1
'                                    vasWork.Value = "1"
''                                    .Row = sList: .Col = 1: .ForeColor = vbBlack
''                                                  .Col = 2: .ForeColor = vbBlack
''                                                  .Col = 3: .ForeColor = vbBlack
''                                                  .Col = 4: .ForeColor = vbBlack
''                                                  .Col = 5: .ForeColor = vbBlack
''                                                  .Col = 6: .ForeColor = vbBlack
''                                                  .Col = 7: .ForeColor = vbBlack
''                                                  .Col = 8: .ForeColor = vbBlack
''                                                  .Col = 9: .ForeColor = vbBlack
''                                                  .Col = 10: .ForeColor = vbBlack
'                                End If
'
'                            Next
'                        End With
'
'                              SQL = "Select ChartNo from pat_res "
'                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
'                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
'                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
'                        res = db_select_Col(gLocal, SQL)
'                        If res = 0 Then
'                                  SQL = " insert into pat_res("
'                            SQL = SQL & "Company,HospCode,ChartNo, "
'                            SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
'                            SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
'                            SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag)"
'                            SQL = SQL & " values ("
'                            SQL = SQL & "'" & XMLInData.Company & "',"
'                            SQL = SQL & "'" & XMLInData.HospCode & "',"
'                            SQL = SQL & "'" & XMLInData.ChartNo & "',"
'                            SQL = SQL & "'" & XMLInData.PatName & "',"
'                            SQL = SQL & "'" & gPatGen.Sex & "',"
'                            SQL = SQL & "'" & gPatGen.Age & "',"
'                            SQL = SQL & "'" & XMLInData.PatJumin & "',"
'                            SQL = SQL & "'" & XMLInData.PatNo & "',"
'                            SQL = SQL & "'" & XMLInData.CommDate & "',"
'                            SQL = SQL & "'" & XMLInData.ExamNo & "',"
'                            SQL = SQL & "'" & XMLInData.ExamID & "',"
'                            SQL = SQL & "'" & XMLInData.ComExamID & "',"
'                            SQL = SQL & "'" & XMLInData.Specimen & "',"
'                            SQL = SQL & "'" & XMLInData.Result & "',"
'                            SQL = SQL & "'" & XMLInData.Reference & "',"
'                            SQL = SQL & "'" & XMLInData.Remark & "',"
'                            SQL = SQL & "'" & XMLInData.RsltDate & "',"
'                            SQL = SQL & "'" & XMLInData.IOFlag & "')"
'                            res = SendQuery(gLocal, SQL)
'                            If res = -1 Then
'                                SaveQuery SQL
'                            End If
'
'                        '-- 속도향상을 위해 쿼리문 지우기
'                        'Else
'                        '          SQL = " Update pat_res Set "
'                        '    SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
'                        '    SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
'                        '    SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
'                        '    SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
'                        '    SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
'                        '    res = SendQuery(gLocal, SQL)
'                        End If
'
'
'                        strBarno = XMLInData.ChartNo
'                    End If
'                End If
'            Next
'            '.Visible = True
'        End With
'    End If
'
'    Exit Sub
'
'    ClearSpread vasList
'    XmlTxt = ""
'    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
'
'    If FindFile <> "" Then
'        FilNum = FreeFile
'        Open "C:\UBCare\SINAI\IF\ExamIF_In.xml" For Input As FilNum
''        Open "\\192.168.0.47\C\UBCare\SINAI\IF\ExamIF_In.xml" For Input As FilNum
'        Do While Not EOF(FilNum)
'            Input #FilNum, TxtString
'            XmlTxt = XmlTxt & TxtString
'        Loop
'        Close #FilNum
'        XmlTxtHead = ""
'        XmlTxtTail = ""
'        TxtPat = TxtString
'        i = InStr(1, TxtPat, "<검사>")
'        X = InStr(1, XmlTxt, "<UBCare검사정보>")
'        XmlTxtHead = Mid(XmlTxt, 1, X + 11)
'        X = InStr(1, XmlTxt, "</UBCare검사정보>")
'        XmlTxtTail = Mid(XmlTxt, X, 13)
'
'        While i > 0
'            If InStr(1, TxtPat, "<검사>") Then
'            '환자별로 text 를 구별
'                i = InStr(1, TxtPat, "<검사>")
'
'                TxtPat = Mid(TxtPat, i)
'                i = InStr(1, TxtPat, "</검사>")
'                TxtRece = Mid(TxtPat, 1, i + 4)
'                TxtPat = Mid(TxtPat, i + 5)
'
'            '차트번호, 환자이름, 주민번호, 내원번호, 검사코드 구분,의뢰일
'            '차트번호
'                i = InStr(1, TxtRece, "<차트번호>")
'                STxt = i + 6
'                i = InStr(1, TxtRece, "</차트번호>")
'                NumTxt = i - STxt
'                PChartNum = Mid(TxtRece, STxt, NumTxt)
'            '환자이름
'                i = InStr(1, TxtRece, "<수진자명>")
'                STxt = i + 6
'                i = InStr(1, TxtRece, "</수진자명>")
'                NumTxt = i - STxt
'                PName = Mid(TxtRece, STxt, NumTxt)
'            '주민번호
'                i = InStr(1, TxtRece, "<주민등록번호>")
'                STxt = i + 8
'                i = InStr(1, TxtRece, "</주민등록번호>")
'                NumTxt = i - STxt
'                PJumin = Mid(TxtRece, STxt, NumTxt)
'                PJumin = Left(PJumin, 6) & Right(PJumin, 7)
'                CalAgeSex PJumin, Format(Date, "yyyy/mm/dd")
'                pSex = gPatGen.Sex
'                PAge = gPatGen.Age
'
'            '내원번호
'                i = InStr(1, TxtRece, "<내원번호>")
'                STxt = i + 6
'                i = InStr(1, TxtRece, "</내원번호>")
'                NumTxt = i - STxt
'                PID = Mid(TxtRece, STxt, NumTxt)
'            '검사코드
'                i = InStr(1, TxtRece, "<검사ID>")
'                STxt = i + 6
'                i = InStr(1, TxtRece, "</검사ID>")
'                NumTxt = i - STxt
'                PExamCode = Mid(TxtRece, STxt, NumTxt)
'            '접수날짜
'                i = InStr(1, TxtRece, "<의뢰일>")
'                STxt = i + 5
'                i = InStr(1, TxtRece, "</의뢰일>")
'                NumTxt = i - STxt
'                PReceDate = Mid(TxtRece, STxt, NumTxt)
'            '<검사번호><입원외래구분>TestNum, IOGubun
'                i = InStr(1, TxtRece, "<검사번호>")
'                STxt = i + 6
'                i = InStr(1, TxtRece, "</검사번호>")
'                NumTxt = i - STxt
'                TestNum = Mid(TxtRece, STxt, NumTxt)
'
'                i = InStr(1, TxtRece, "<입원외래구분>")
'                STxt = i + 8
'                i = InStr(1, TxtRece, "</입원외래구분>")
'                NumTxt = i - STxt
'                IOGubun = Mid(TxtRece, STxt, NumTxt)
'
'                    SQL = "select equipno, equipcode, examname from equipexam where examcode = '" & PExamCode & "' "
'                    res = db_select_Col(gLocal, SQL)
'
'                    If res > 0 Then
'
'                        PEquipno = gReadBuf(0)
'                        PEquipCode = gReadBuf(1)
'                        PExamname = gReadBuf(2)
'
'                        SQL = "select barcode from pat_res where barcode = '" & PChartNum & "' and examcode = '" & PExamCode & "' and recedate = '" & PReceDate & "'"
'                        res = db_select_Col(gLocal, SQL)
'                        If res = 0 Then
'                        SQL = "insert into pat_res(equipno, barcode,equipcode, " & vbCrLf & _
'                              "examcode, recedate, pid,pname, psex, page, pjumin, examname,examdate, gubun, subcode, result) " & vbCrLf & _
'                              "values('" & PEquipno & "','" & PChartNum & "'," & vbCrLf & _
'                              "'" & PEquipCode & "','" & PExamCode & "'," & vbCrLf & _
'                              "'" & PReceDate & "','" & PID & "'," & vbCrLf & _
'                              "'" & PName & "','" & pSex & "'," & vbCrLf & _
'                              "'" & PAge & "','" & PJumin & "'," & vbCrLf & _
'                              "'" & PExamname & "','" & Format(Date, "yyyymmdd") & "','" & IOGubun & "','" & TestNum & "','')"
'                              res = SendQuery(gLocal, SQL)
'                              If res = -1 Then
'                                SaveQuery SQL
'                              End If
'
'                        Else
'                            SQL = " Update pat_res Set " & CR & _
'                                  " barcode = '" & PChartNum & "', " & CR & _
'                                  " pname = '" & PName & "', " & CR & _
'                                  " psex = '" & pSex & "', " & CR & _
'                                  " page = '" & PAge & "', " & CR & _
'                                  " recedate =  '" & PReceDate & "', " & CR & _
'                                  " pjumin =  '" & PJumin & "', " & CR & _
'                                  " subcode =  '" & TestNum & "', " & CR & _
'                                  " gubun =  '" & IOGubun & "', " & CR & _
'                                  " pid =  '" & PID & "' " & CR & _
'                                  " Where examdate = '" & Format(frmInterface.txtToday.Text, "yyyymmdd") & "' " & CR & _
'                                  " and examcode = '" & PExamCode & "'" & CR & _
'                                  " and barcode = '" & PChartNum & "' "
'                            res = SendQuery(gLocal, SQL)
'                        End If
'
'
'                        BarFlag = 0
'                        For j = 1 To vasList.DataRowCnt
'                            If GetText(vasList, j, 11) = PChartNum Then
'                                BarFlag = 1
'                            End If
'                        Next
'    '                    If BarFlag = 0 Then
'    '                        SQL = " Select pid,  pname,  psex, page,pjumin,recedate, barcode " & CR & _
'    '                              " From pat_res " & CR & _
'    '                              " Where barcode = '" & PChartNum & "' and pid = '" & PID & "' and recedate = '" & PReceDate & "' " & CR & _
'    '                              " Group By barcode, pname, psex, page, pjumin, recedate, pid "
'    '                        res = db_select_Vas(gLocal, SQL, vasList, vasList.DataRowCnt + 1, 5)
'    '                    End If
'
'                    End If
'
'                i = InStr(1, TxtPat, "<검사>")
'            Else
'            End If
'
'        Wend
'        XmlTxtTail = TxtPat
'        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"
'        ClearSpread vasList
'        SQL = " Select barcode,  pname,  psex, page,pjumin,recedate, pid, '', gubun " & CR & _
'              " From pat_res " & CR & _
'              " Where  recedate >= '" & sSch1 & "' and recedate <= '" & sSch2 & "' and result = '' " & CR & _
'              " Group By barcode, pname, psex, page, pjumin, recedate, pid ,gubun"
'        res = db_select_Vas(gLocal, SQL, vasList, 1, 2)
'
'        vasList.SetText 12, iRow, "A" & Format(Trim(GetText(vasList, iRow, 9)), "yyyymmdd") & "-" & Trim(GetText(vasList, iRow, 10))
'
'    Else
'        ClearSpread vasList
'        SQL = " Select barcode,  pname,  psex, page,pjumin,recedate, pid, '',gubun " & CR & _
'              " From pat_res " & CR & _
'              " Where  recedate >= '" & sSch1 & "' and recedate <= '" & sSch2 & "' and result = '' " & CR & _
'              " Group By barcode, pname, psex, page, pjumin, recedate, pid,gubun "
'        res = db_select_Vas(gLocal, SQL, vasList, 1, 4)
'
'        vasList.SetText 12, iRow, "O" & Format(Trim(GetText(vasList, iRow, 9)), "yyyymmdd") & "-" & Trim(GetText(vasList, iRow, 10))
'     End If
'
'    vasList.MaxRows = vasList.DataRowCnt
'    vasList.RowHeight(-1) = 13.3
'
'End Sub
Private Sub cmdSetup_Click()
'    frmEquipExam.SSPanel1.Caption = "  BT2000 plus 장비 코드 설정"
    frmEquipExam.SSPanel1.Caption = "  URIT8021A 장비 코드 설정"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub cmdWorkList_Click()
    frmPatSear.Left = 0
    frmPatSear.Top = 0
    frmPatSear.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    txtMsg.Text = ""
    
    vasRes.OperationMode = 0
    
'    vasID.MaxRows = 0
    vasID.SetFocus
    
    ClearSpread vasRes, 1, 1
    
    vasRes.MaxRows = 0
    
    vasID.RowHeight(-1) = 14
    
End Sub

Private Sub Form_Load()
    Dim sDate As String
    '1. 화면 및 변수 초기화
    '2. 데이타베이스에 Connect 하기 - Local - Server
    '3. Ini 내용 불러오기    GetSetup
    '4. Comport Open
    
    Dim i As Integer
    
    For i = 0 To 9
        txtMain(i).Text = ""
    Next
    
    For i = 0 To 2
        txtMain1(i).Text = ""
    Next

    Me.Left = 0
    Me.Top = 0
    
    'Clear
    txtMsg.Text = ""

    ClearSpread vasID
'    ClearSpread vasWork
'    ClearSpread vasList
    
    'vasActiveCell vasID, 1, colPID
    
    vasRes.OperationMode = 0
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 1
    
    GetSetup    'ini에서 DB정보 불러오기
        
    'If Not Connect_Server Then
    '    MsgBox "서버에 연결되지 않았습니다."
        'Exit Sub
    'End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        'Exit Sub
    End If

'    MSComm1.CommPort = 3
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
'    MSComm1.Settings = "9600,n,7,2"
'
'    MSComm2.CommPort = gSetup.gPort
'    MSComm2.RTSEnable = gSetup.gRTSEnable
'    MSComm2.DTREnable = gSetup.gDTREnable
'    MSComm2.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    Me.txtUID = gExamUID
    
    raw_data = ""
    
'    If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'    End If
    
'    If MSComm2.PortOpen = False Then
'        MSComm2.PortOpen = True
'    End If
    
    txtToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpDate = Format(CDate(GetDateFull), "yyyy/mm/dd")
    
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("d", CDate(txtToday.Text), -30), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '검사코드 가져오기
    GetExamCode
        
    'MultiSelect Mode
    vasRes.OperationMode = 1
    
    'fontsize
    vasRes.FontSize = 9
    
    'dtpStartDt.Value = Now - 3
    'dtpStopDt.Value = Now

    FileURIT.Path = gMachPath

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    blnSameRecord = False
    
    vasID.RowHeight(-1) = 12
    vasRes.RowHeight(-1) = 12

End Sub

''Private Sub SetPatInfo(ByVal pBarNo As String, Optional bRow As Long)
''    Dim i           As Integer
''    Dim intRow      As Long
''    Dim strItems    As String
''    Dim strGbn      As String
''    Dim varTmp      As Variant
''
''    intRow = -1
''    For i = bRow To vasID.DataRowCnt
''        If Trim(GetText(vasID, i, colPID)) = pBarNo Then
''            intRow = i
''            strGbn = vasID.GetText(3, intRow, varTmp)
''            strGbn = varTmp
''            Exit For
''        End If
''    Next i
''
''    If intRow < 0 Then
'''        intRow = vasID.DataRowCnt + 1
'''        If vasID.MaxRows < intRow Then
'''            vasID.MaxRows = intRow
'''        End If
''        gRow = -1
''        Exit Sub
''    End If
''
''    '-- 장비수신정보 표시
''    'Call SetText(vasID, pBarNo, intRow, colBarcode)             '2 Barcode
''    'Call SetText(vasID, mResult.RackNo, intRow, colRack)        '3 Rack
''    'Call SetText(vasID, mResult.TubePos, intRow, colPos)        '4 Pos
''
''    Call SetText(vasID, "Result", intRow, colRCnt)    '상태
''
''    Call vasActiveCell(vasID, intRow, colBarcode)
''
''    '-- 결과스프레드 지우기
''    Call ClearSpread(vasRes)
''
''    '-- 검사자 정보 서버테이블 가져와 표시(for 워크리스트)  '5,6,7,8
''    'Call GetSampleInfoW(intRow)                                '5,6,7,8
''
''    '-- 현재 Row
''    gRow = intRow
''
''    '-- 바코드번호에 존재하는 검사코드 가져오기(인수 : 장비코드,바코드번호)
''    gOrderExam = GetOrderExamCode(gEquip, pBarNo, intRow)
''
''End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
'''Private Sub EditRcvDataURIT()
'''    Dim strRcvBuf    As String   '수신한 Data
'''    Dim strType      As String   '수신한 Record Type
'''    Dim strBarno     As String   '수신한 바코드번호
'''    Dim strSeq       As String   '수신한 Sequence
'''    Dim strRackNo    As String   '수신한 Rack Or Disk No
'''    Dim strTubePos   As String   '수신한 Tube Position
'''    Dim strIntBase   As String   '수신한 장비기준 검사명
'''    Dim strResult    As String   '수신한 결과
'''    Dim strQCResult  As String   '수신한 결과(QC)
'''    Dim strFlag      As String   '수신한 Abnormal Flag
'''    Dim strComm      As String   '수신한 Comment
'''    Dim strTemp1     As String
'''    Dim strTemp2     As String
'''    Dim intCnt       As Integer
'''
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResult_Buff As String
'''    Dim lsExamDate As String
'''    Dim lsEquipRes As String
'''    Dim lsResRow    As String
'''    Dim ii As Integer
'''    Dim strTmp      As String
'''    Dim intIdx      As Integer
'''    Dim varTmp      As Variant
'''    Dim intgRow     As Long
'''
'''    For intCnt = 1 To UBound(strRecvData)
'''        strRcvBuf = strRecvData(intCnt)
'''        strType = Mid$(strRcvBuf, 1, 2)
'''
'''        If Trim(Mid(strRcvBuf, 13, 12)) <> "" Then
'''            strBarno = Trim(Mid(strRcvBuf, 13, 12))
'''
'''            With mResult
'''                .BarNo = strBarno
'''            End With
'''
'''            strTmp = Mid$(strRcvBuf, 108)
'''            intgRow = 1
'''
'''            Call SetPatInfo(strBarno, intgRow)
'''            Call ClearSpread(vasRes)
'''
'''            If gRow > 1 Then
'''
'''                varTmp = Split(strTmp, ";")
'''
'''                For ii = 0 To UBound(varTmp)
'''                    strIntBase = mGetP(varTmp(ii), 1, "=")
'''                    strResult = mGetP(varTmp(ii), 2, "=")
'''
'''                    If strResult <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                        SQL = SQL & "  FROM EQPMASTER"
'''                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''
'''                        res = GetDBSelectColumn(gLocal, SQL)
'''
'''                        '-- 오더 있을 경우
'''                        If res > 0 Then
'''                            lsExamCode = Trim(gReadBuf(0))
'''                            lsExamName = Trim(gReadBuf(1))
'''                            lsSeqNo = Trim(gReadBuf(2))
'''
'''                            lsResRow = vasRes.DataRowCnt + 1
'''                            If vasRes.MaxRows < lsResRow Then
'''                                vasRes.MaxRows = lsResRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            lsEquipRes = strResult
'''                            strResult = SetResult(strResult, strIntBase)
'''                            lsResult_Buff = strResult
'''
'''                            '-- Work List
'''                            SetText vasID, "Result", gRow, colRCnt                 '10 진행상태
'''                            vasID.Row = gRow
'''                            vasID.Row2 = gRow
'''                            vasID.Col = 2
'''                            vasID.Col2 = vasID.MaxCols
'''                            vasID.BlockMode = True
'''                            vasID.BackColor = vbCyan
'''                            vasID.BlockMode = False
'''
'''                            '-- 결과 List
'''                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
'''                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
'''                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
'''                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'''                            SetText vasRes, strResult, lsResRow, colResult          '결과
'''                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                            lsResult_Buff = ""
'''
'''                            strState = "R"
'''
'''                        End If
'''                    End If
'''                    'strTmp = Mid$(strTmp, 12)
'''                Next
'''
'''
''''''                If MnTransAuto.Checked = True And strState = "R" Then
''''''
''''''                    res = SaveTransDataW(gRow)
''''''
''''''                    If res = -1 Then
''''''                        '-- 저장 실패
''''''                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
''''''                        SetText vasID, "Failed", gRow, colState
''''''                    Else
''''''                        '-- 저장 성공
''''''                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
''''''                        SetText vasID, "Trans", gRow, colState
''''''
''''''                        SQL = " Update PATRESULT Set " & vbCrLf & _
''''''                              " sendflag = '2' " & vbCrLf & _
''''''                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
''''''                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
''''''                        res = SendQuery(gLocal, SQL)
''''''                        If res = -1 Then
''''''                            SaveQuery SQL
''''''                            Exit Sub
''''''                        End If
''''''                    End If
''''''                End If
'''
'''                'SetText vasID, "Result", gRow, colState
'''                strState = ""
'''            End If
'''
'''            '===================================================================
'''            Call SetPatInfo(strBarno, gRow + 1)
'''
'''            If gRow > 1 Then
'''                varTmp = Split(strTmp, ";")
'''
'''                For ii = 0 To UBound(varTmp)
'''                    strIntBase = mGetP(varTmp(ii), 1, "=")
'''                    strResult = mGetP(varTmp(ii), 2, "=")
'''                    'strComm = Mid$(strTmp, 10, 1)
'''
'''                    If strResult <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                        SQL = SQL & "  FROM EQPMASTER"
'''                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''
'''                        res = GetDBSelectColumn(gLocal, SQL)
'''
'''                        '-- 오더 있을 경우
'''                        If res > 0 Then
'''                            lsExamCode = Trim(gReadBuf(0))
'''                            lsExamName = Trim(gReadBuf(1))
'''                            lsSeqNo = Trim(gReadBuf(2))
'''
'''                            lsResRow = vasRes.DataRowCnt + 1
'''                            If vasRes.MaxRows < lsResRow Then
'''                                vasRes.MaxRows = lsResRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            lsEquipRes = strResult
'''                            strResult = SetResult(strResult, strIntBase)
'''                            lsResult_Buff = strResult
'''
'''                            '-- Work List
'''                            SetText vasID, "Result", gRow, colRCnt                 '10 진행상태
'''                            vasID.Row = gRow
'''                            vasID.Row2 = gRow
'''                            vasID.Col = 2
'''                            vasID.Col2 = vasID.MaxCols
'''                            vasID.BlockMode = True
'''                            vasID.BackColor = vbCyan
'''                            vasID.BlockMode = False
'''
'''                            '-- 결과 List
'''                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
'''                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
'''                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
'''                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'''                            SetText vasRes, strResult, lsResRow, colResult          '결과
'''                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                            lsResult_Buff = ""
'''
'''                            strState = "R"
'''
'''                        End If
'''                    End If
'''                Next
'''
'''
''''''                If MnTransAuto.Checked = True And strState = "R" Then
''''''
''''''                    res = SaveTransDataW(gRow)
''''''
''''''                    If res = -1 Then
''''''                        '-- 저장 실패
''''''                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
''''''                        SetText vasID, "Failed", gRow, colState
''''''                    Else
''''''                        '-- 저장 성공
''''''                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
''''''                        SetText vasID, "Trans", gRow, colState
''''''
''''''                        SQL = " Update PATRESULT Set " & vbCrLf & _
''''''                              " sendflag = '2' " & vbCrLf & _
''''''                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
''''''                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
''''''                        res = SendQuery(gLocal, SQL)
''''''                        If res = -1 Then
''''''                            SaveQuery SQL
''''''                            Exit Sub
''''''                        End If
''''''                    End If
''''''                End If
'''
'''
'''                'SetText vasID, "Result", gRow, colState
'''                strState = ""
'''            End If
'''
'''            '===================================================================
'''            Call SetPatInfo(strBarno, gRow + 1)
'''
'''            If gRow > 1 Then
'''                varTmp = Split(strTmp, ";")
'''
'''                For ii = 0 To UBound(varTmp)
'''                    strIntBase = mGetP(varTmp(ii), 1, "=")
'''                    strResult = mGetP(varTmp(ii), 2, "=")
'''                    'strComm = Mid$(strTmp, 10, 1)
'''
'''                    If strResult <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                        SQL = SQL & "  FROM EQPMASTER"
'''                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''
'''                        res = GetDBSelectColumn(gLocal, SQL)
'''
'''                        '-- 오더 있을 경우
'''                        If res > 0 Then
'''                            lsExamCode = Trim(gReadBuf(0))
'''                            lsExamName = Trim(gReadBuf(1))
'''                            lsSeqNo = Trim(gReadBuf(2))
'''
'''                            lsResRow = vasRes.DataRowCnt + 1
'''                            If vasRes.MaxRows < lsResRow Then
'''                                vasRes.MaxRows = lsResRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            lsEquipRes = strResult
'''                            strResult = SetResult(strResult, strIntBase)
'''                            lsResult_Buff = strResult
'''
'''                            '-- Work List
'''                            SetText vasID, "Result", gRow, colRCnt                 '10 진행상태
'''                            vasID.Row = gRow
'''                            vasID.Row2 = gRow
'''                            vasID.Col = 2
'''                            vasID.Col2 = vasID.MaxCols
'''                            vasID.BlockMode = True
'''                            vasID.BackColor = vbCyan
'''                            vasID.BlockMode = False
'''
'''                            '-- 결과 List
'''                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
'''                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
'''                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
'''                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'''                            SetText vasRes, strResult, lsResRow, colResult          '결과
'''                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                            lsResult_Buff = ""
'''
'''                            strState = "R"
'''
'''                        End If
'''                    End If
'''                Next
'''
'''
''''''                If MnTransAuto.Checked = True And strState = "R" Then
''''''
''''''                    res = SaveTransDataW(gRow)
''''''
''''''                    If res = -1 Then
''''''                        '-- 저장 실패
''''''                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
''''''                        SetText vasID, "Failed", gRow, colState
''''''                    Else
''''''                        '-- 저장 성공
''''''                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
''''''                        SetText vasID, "Trans", gRow, colState
''''''
''''''                        SQL = " Update PATRESULT Set " & vbCrLf & _
''''''                              " sendflag = '2' " & vbCrLf & _
''''''                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
''''''                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
''''''                        res = SendQuery(gLocal, SQL)
''''''                        If res = -1 Then
''''''                            SaveQuery SQL
''''''                            Exit Sub
''''''                        End If
''''''                    End If
''''''                End If
''''''
''''''
'''                'SetText vasID, "Result", gRow, colState
'''                strState = ""
'''            End If
'''
'''
'''        End If
'''    Next
'''
'''End Sub
'''
'''' asRow1 = Work List
'''' asRow2 = 결과 List
'''Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
'''    Dim sCnt As String
'''    Dim sExamDate As String
'''
'''    sExamDate = Format(dtpToday, "yyyymmdd")
'''
'''    SQL = ""
'''    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
'''          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'''          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'''          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colPID)) & "' " & vbCrLf & _
'''          "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'''          "  AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'''    SQL = SQL & " AND SAMPLETYPE = '" & Trim(GetText(vasID, asRow1, colDISK)) & "'"
'''
'''    res = SendQuery(gLocal, SQL)
'''
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Function
'''    End If
'''
'''    SQL = ""
'''    SQL = SQL & "INSERT INTO PATRESULT("
'''    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,SAMPLETYPE,DISKNO,POSNO," & vbCrLf & _
'''                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'''                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'''    SQL = SQL & "VALUES("
'''    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'''    SQL = SQL & "'" & gEquip & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
'''    SQL = SQL & "'', "
'''    SQL = SQL & "'', " & vbCrLf
'''    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
'''    SQL = SQL & "'', "
'''    SQL = SQL & "'', "
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colMachResult)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'''    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
'''    SQL = SQL & "'0', "
'''    SQL = SQL & "'" & gIFUser & "')"
'''
'''    res = SendQuery(gLocal, SQL)
'''
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Function
'''    End If
'''
'''End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If MSComm2.PortOpen = True Then MSComm2.PortOpen = False
    'DisConnect_Server
    DisConnect_Local
    
    Unload Me
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    Dim j As Integer
    
    gAllExam = ""
    
    ClearSpread vasTemp
    
    SQL = "Select EquipCode, ExamCode, ExamName From EquipExam where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by EquipCode"
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res > 0 Then
        ReDim gArr_ExamCode(1 To vasTemp.DataRowCnt, 1 To 3)
    Else
        SaveQuery SQL
    End If
    
    For i = 1 To vasTemp.DataRowCnt
        gArr_ExamCode(i, 1) = i
        
        For j = 1 To 2
            gArr_ExamCode(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
    Next i
    
End Sub


Private Sub BT2000_Bak(asData As String)
    Dim i As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lsUnitNo As String
    Dim lsRackNo As String
    Dim lsPos As String
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sPoint As String        '소수점
    Dim sTmpStr As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    
    Dim lsSeqNo As String
    
    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim iLen As String
    
    If Trim(asData) = "" Then
        Exit Sub
    End If
    
    lsPID = Trim(Mid(asData, 1, 15))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colPID)) = lsPID Then
            llRow = iRow
            Exit For
        End If
    Next iRow
     
    If llRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow + 1
        End If
    End If
        
    vasActiveCell vasID, llRow, colPID
         
    ClearSpread vasRes, 1, 1
        
    SetText vasID, lsPID, llRow, colPID
    If Trim(GetText(vasID, llRow, colPName)) = "" Then
        'Get_Sample_Info llRow
    End If
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))
        
    '결과 잘라 넣기
    j = 0
                        
    lsData = Mid(asData, 21)
    Do While Len(lsData) >= 5
        lsCode = Trim(Left(lsData, 4))
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0.0#")
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0#")
        lsRt = Trim(Mid(lsData, 5, 7))
        
        gReadBuf(0) = "0"
        '검사코드, 순서(서브코드), 검사명, 소수점
        SQL = "Select examcode, examname, resprec From equipexam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  And equipcode = '" & lsCode & "'"
        res = db_select_Col(gLocal, SQL)

        If (res = 1) And (gReadBuf(0) <> "") Then
            j = j + 1

            If IsNumeric(lsRt) Then
                sExamCode = Trim(gReadBuf(0))
                sExamName = Trim(gReadBuf(1))
                sPoint = Trim(gReadBuf(2))
                
                '소수점 처리
                If IsNumeric(sPoint) Then
                    If CInt(sPoint) > 0 Then
                        sTmpStr = "#0."
                        For i = 1 To CInt(sPoint)
                            sTmpStr = sTmpStr & "0"
                        Next i
                    Else
                        sTmpStr = "#0"
                    End If
                    
                    lsRt = Format(lsRt, sTmpStr)
                End If
                
                SetText vasRes, Trim(GetText(vasID, llRow, colPID)), j, 2  '검체번호

                SetText vasRes, lsCode, j, colEquipExam '장비코드
                SetText vasRes, sExamCode, j, colExamCode   '검사코드
                SetText vasRes, sExamName, j, colExamName   '검사명
                SetText vasRes, lsRt, j, colResult          '검사결과
                SetText vasRes, lsRt, j, colResult1         '검사결과
                SetText vasRes, lsFlag, j, colRCheck        '판정
                
                '로컬
                Save_Local_One llRow, j, "A"
            End If
        End If
        
        lsData = Mid(lsData, 12)
    Loop

    gReadBuf(0) = ""
    
    '수신완료======================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    
End Sub

Private Sub BT2000(asData As String)
    
    Dim i As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lsUnitNo As String
    Dim lsRackNo As String
    Dim lsPos As String
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sPoint As String        '소수점
    Dim sTmpStr As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    
    Dim lsSeqNo As String
    
    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim iLen As String
    
    Dim lsEquip As String
    
'    lsEquip = "BT2000"
    lsEquip = "URIT8021A"
    
    lsPID = Trim(Mid(asData, 1, 15))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colPID)) = lsPID Then
            llRow = iRow
            Exit For
        End If
    Next iRow
     
    If llRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow + 1
        End If
    End If
        
    vasActiveCell vasID, llRow, colPID
         
    ClearSpread vasRes, 1, 1
        
    SetText vasID, lsPID, llRow, colPID
'    If Trim(GetText(vasID, llRow, colPName)) = "" Then
'        Get_Sample_Info llRow
'    End If
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & lsEquip & "' "
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))
        
    '결과 잘라 넣기
    j = 0
                        
    lsData = Mid(asData, 21)
    Do While Len(lsData) >= 5
        lsCode = Trim(Left(lsData, 4))
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0.0#")
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0#")
        lsRt = Trim(Mid(lsData, 5, 7))
        
        gReadBuf(0) = "0"
        '검사코드, 순서(서브코드), 검사명, 소수점
        SQL = "Select examcode, examname, resprec From equipexam" & vbCrLf & _
              " Where Equipno = '" & lsEquip & "' " & vbCrLf & _
              "  And equipcode = '" & lsCode & "'"
        res = db_select_Col(gLocal, SQL)

        If (res = 1) And (gReadBuf(0) <> "") Then
            j = j + 1

            If IsNumeric(lsRt) Then
                sExamCode = Trim(gReadBuf(0))
                sExamName = Trim(gReadBuf(1))
                sPoint = Trim(gReadBuf(2))
                
                '소수점 처리
                If IsNumeric(sPoint) Then
                    If CInt(sPoint) > 0 Then
                        sTmpStr = "#0."
                        For i = 1 To CInt(sPoint)
                            sTmpStr = sTmpStr & "0"
                        Next i
                    Else
                        sTmpStr = "#0"
                    End If
                    
                    lsRt = Format(lsRt, sTmpStr)
                End If
                
                SetText vasRes, Trim(GetText(vasID, llRow, colPID)), j, 2  '검체번호

                SetText vasRes, lsCode, j, colEquipExam '장비코드
                SetText vasRes, sExamCode, j, colExamCode   '검사코드
                SetText vasRes, sExamName, j, colExamName   '검사명
                SetText vasRes, lsRt, j, colResult          '검사결과
                SetText vasRes, lsRt, j, colResult1         '검사결과
                SetText vasRes, lsFlag, j, colRCheck        '판정
                
                '로컬
                Save_Local_One llRow, j, lsEquip
            End If
        End If
        
        lsData = Mid(lsData, 12)
    Loop

    gReadBuf(0) = ""
    
    '수신완료======================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    
End Sub

Private Sub URIT8021A(asData As Variant)
    
    Dim i As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lsUnitNo As String
    Dim lsRackNo As String
    Dim lsPos As String
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sPoint As String        '소수점
    Dim sTmpStr As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    
    Dim lsSeqNo As String
    
    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim iLen As String
    
    Dim lsEquip As String
    
    Dim intCnt      As Integer
    Dim strRcvBuf   As String
    Dim strType     As String
    Dim strTmp      As String
    Dim ii          As Long
    Dim varTmp      As Variant
    
    lsEquip = "URIT8021A"
   
    For intCnt = 0 To UBound(asData)
        strRcvBuf = asData(intCnt)
        strType = Mid$(strRcvBuf, 1, 2)
        
        If Trim(Mid(strRcvBuf, 13, 12)) <> "" Then
            lsPID = Trim(Mid(strRcvBuf, 13, 12))
        End If
        
        'strTmp = Mid$(strRcvBuf, 108)
        'intgRow = 1
    
        'lsPID = "16453"
        
        '같은 바코드번호의 검체는 디스플레이되지 않음
        llRow = -1
        For iRow = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, iRow, colPID)) = lsPID Then
                llRow = iRow
                Exit For
            End If
        Next iRow
         
        If llRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            llRow = vasID.DataRowCnt + 1
            If llRow > vasID.MaxRows Then
                vasID.MaxRows = llRow + 1
            End If
        End If
            
        vasActiveCell vasID, llRow, colPID
             
        ClearSpread vasRes, 1, 1
            
        SetText vasID, lsPID, llRow, colPID
        
        '수신중========================================================
        SetText vasID, "수신중", llRow, colState
        SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
        '==============================================================
        
        '검사코드만큼 Row의 갯수를 설정
        gReadBuf(0) = "0"
        
        SQL = "Select count(examcode) From equipexam" & vbCrLf & _
              " Where equipno = '" & lsEquip & "' "
        res = db_select_Col(gLocal, SQL)
    
        vasRes.MaxRows = Trim(gReadBuf(0))
            
''                varTmp = Split(strTmp, ";")
''
''                For ii = 0 To UBound(varTmp)
''                    strIntBase = mGetP(varTmp(ii), 1, "=")
''                    strResult = mGetP(varTmp(ii), 2, "=")
''
''                    If strResult <> "" Then
        
        '결과 잘라 넣기
        j = 0
                            
        lsData = Mid(strRcvBuf, 108)
        varTmp = Split(lsData, ";")
                
        For ii = 0 To UBound(varTmp)
'            strIntBase = mGetP(varTmp(ii), 1, "=")
'            strResult = mGetP(varTmp(ii), 2, "=")
            lsCode = mGetP(varTmp(ii), 1, "=")
            lsRt = mGetP(varTmp(ii), 2, "=")
            
            gReadBuf(0) = "0"
            '검사코드, 순서(서브코드), 검사명, 소수점
            SQL = "Select examcode, examname, resprec From equipexam" & vbCrLf & _
                  " Where Equipno = '" & lsEquip & "' " & vbCrLf & _
                  "  And equipcode = '" & lsCode & "'"
            res = db_select_Col(gLocal, SQL)
    
            If (res = 1) And (gReadBuf(0) <> "") Then
                j = j + 1
    
                If IsNumeric(lsRt) Then
                    sExamCode = Trim(gReadBuf(0))
                    sExamName = Trim(gReadBuf(1))
                    sPoint = Trim(gReadBuf(2))
                    
                    '소수점 처리
                    If IsNumeric(sPoint) Then
                        If CInt(sPoint) > 0 Then
                            sTmpStr = "#0."
                            For i = 1 To CInt(sPoint)
                                sTmpStr = sTmpStr & "0"
                            Next i
                        Else
                            sTmpStr = "#0"
                        End If
                        
                        lsRt = Format(lsRt, sTmpStr)
                    End If
                    
                    SetText vasRes, Trim(GetText(vasID, llRow, colPID)), j, 2  '검체번호
    
                    SetText vasRes, lsCode, j, colEquipExam '장비코드
                    SetText vasRes, sExamCode, j, colExamCode   '검사코드
                    SetText vasRes, sExamName, j, colExamName   '검사명
                    SetText vasRes, lsRt, j, colResult          '검사결과
                    SetText vasRes, lsRt, j, colResult1         '검사결과
                    SetText vasRes, lsFlag, j, colRCheck        '판정
                    
                    '로컬
                    Save_Local_One llRow, j, lsEquip
                End If
            End If
            
            'lsData = Mid(lsData, 12)
        Next
    
        gReadBuf(0) = ""
    
    Next
    
    '수신완료======================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    
End Sub

Private Sub BC2800(asData As String)
    
    Dim i As Integer
    Dim j As Integer
    Dim iArry As Integer
    Dim iPoint As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lsUnitNo As String
    Dim lsRackNo As String
    Dim lsPos As String
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sExamName As String
    Dim sPoint As String        '소수점
    Dim sTmpStr As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim IsRt1(20) As String
    Dim lsFlag As String
    
    Dim lsSeqNo As String
    
    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim sTmpVal As String
    Dim sTmpSeq As String
    
    Dim iLen As String
    
    Dim lsEquip As String
    
    
    lsEquip = "BC2800"
    
    If Trim(asData) = "" Then
        Exit Sub
    End If
    
    lsPID = CLng(Trim(Mid(asData, 4, 6)))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colPID)) = CLng(lsPID) Then
            llRow = iRow
            Exit For
        End If
    Next iRow
     
    If llRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow + 1
        End If
    End If
        
    vasActiveCell vasID, llRow, colPID
         
    ClearSpread vasRes, 1, 1
        
    SetText vasID, lsPID, llRow, colPID
'    If Trim(GetText(vasID, llRow, colPName)) = "" Then
'        Get_Sample_Info llRow
'    End If
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & lsEquip & "' "
    
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))
        
    '결과 잘라 넣기
    j = 0
    iArry = 0
    
    lsData = Mid(asData, 23, 66)
    
    '0088 0031 0009 00483571045393911160332089402961603490053097163051049
   ' 0088 0031 0009 0048 357 104 539 391 116 0332 0894 0296 160 349 0053 097 163 051 0499

    IsRt1(1) = Left(lsData, 3) & "." & Mid(lsData, 4, 1) & " 01"
    IsRt1(2) = Mid(lsData, 5, 3) & "." & Mid(lsData, 8, 1) & " 02"
    IsRt1(3) = Mid(lsData, 9, 3) & "." & Mid(lsData, 12, 1) & " 03"
    IsRt1(4) = Mid(lsData, 13, 3) & "." & Mid(lsData, 16, 1) & " 04"
    IsRt1(5) = Mid(lsData, 17, 2) & "." & Mid(lsData, 19, 1) & " 05"
    IsRt1(6) = Mid(lsData, 20, 2) & "." & Mid(lsData, 22, 1) & " 06"
    IsRt1(7) = Mid(lsData, 23, 2) & "." & Mid(lsData, 25, 1) & " 07"
    IsRt1(8) = Mid(lsData, 26, 1) & "." & Mid(lsData, 27, 2) & " 08"
    IsRt1(9) = Mid(lsData, 29, 2) & "." & Mid(lsData, 31, 1) & " 09"
    IsRt1(10) = Mid(lsData, 32, 3) & "." & Mid(lsData, 35, 1) & " 10"
    IsRt1(11) = Mid(lsData, 36, 3) & "." & Mid(lsData, 39, 1) & " 11"
    IsRt1(12) = Mid(lsData, 40, 3) & "." & Mid(lsData, 43, 1) & " 12"
    IsRt1(13) = Mid(lsData, 44, 2) & "." & Mid(lsData, 46, 1) & " 13"
    IsRt1(14) = Mid(lsData, 47, 2) & "." & Mid(lsData, 49, 1) & " 14"
    IsRt1(15) = Mid(lsData, 50, 4) & "  15"
    IsRt1(16) = Mid(lsData, 54, 2) & "." & Mid(lsData, 56, 1) & " 16"
    IsRt1(17) = Mid(lsData, 57, 2) & "." & Mid(lsData, 59, 1) & " 17"
    IsRt1(18) = Mid(lsData, 60, 1) & "." & Mid(lsData, 61, 3) & " 18"
    IsRt1(19) = Mid(lsData, 64, 2) & "." & Mid(lsData, 66, 1) & " 19"

    Do Until iArry = 20
        'lsCode = Trim(Left(lsData, 4))
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0.0#")
        'lsRt = Format(Trim(Mid(lsData, 5, 7)), "0#")
        'lsRt = Trim(Mid(lsData, 5, 7))
            
        iArry = iArry + 1
        sTmpVal = Left(Trim(IsRt1(iArry)), 5)
        sTmpSeq = Right(IsRt1(iArry), 2)
        
        If IsNumeric(sTmpVal) = False Then
            sTmpVal = "0"
        End If
        
        gReadBuf(0) = "0"
        '검사코드, 순서(서브코드), 검사명, 소수점
        SQL = "Select equipcode, examcode, examname, resprec From equipexam" & vbCrLf & _
              " Where Equipno = '" & lsEquip & "' And deltavalue = '" & sTmpSeq & "' "
             
        res = db_select_Col(gLocal, SQL)

        If (res = 1) And (gReadBuf(0) <> "") Then
            j = j + 1
            
            If IsNumeric(sTmpVal) Then
                lsCode = Trim(gReadBuf(0))
                sExamCode = Trim(gReadBuf(1))
                sExamName = Trim(gReadBuf(2))
                sPoint = Trim(gReadBuf(3))
                
                '소수점 처리
                If IsNumeric(sPoint) Then
                    If CInt(sPoint) > 0 Then
                        sTmpStr = "#0."
                        For i = 1 To CInt(sPoint)
                            sTmpStr = sTmpStr & "0"
                        Next i
                    Else
                        sTmpStr = "#0"
                    End If
                    
                    sTmpVal = Format(sTmpVal, sTmpStr)
                End If
                
                SetText vasRes, Trim(GetText(vasID, llRow, colPID)), j, 2  '검체번호

                SetText vasRes, lsCode, j, colEquipExam '장비코드
                SetText vasRes, sExamCode, j, colExamCode   '검사코드
                SetText vasRes, sExamName, j, colExamName   '검사명
                SetText vasRes, sTmpVal, j, colResult          '검사결과
                SetText vasRes, lsRt, j, colResult1         '검사결과
                SetText vasRes, lsFlag, j, colRCheck        '판정
                
                '로컬
                Save_Local_One_CBC llRow, j, lsEquip
            End If
        End If
        
        'lsData = Mid(lsData, 12)
    Loop

    gReadBuf(0) = ""
    
    '수신완료======================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsPID As String
    Dim lsReceNo As String
    
    '샘플 환자 정보 가져오기
    lsPID = Trim(GetText(vasID, asRow, colPID))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이, 처방개수
    SQL = " Select PbsPatNam, PbsSexTyp, PbsBirDte " & vbCrLf & _
          " From PbsInf " & vbCrLf & _
          " Where PbsChtNum = " & lsPID & ""
    res = db_select_Col(gServer, SQL)
    
    If gReadBuf(0) <> "" Then
        SetText vasID, gReadBuf(0), asRow, colPName
        SetText vasID, gReadBuf(1), asRow, colPSex

        CalAgeSex gReadBuf(2), txtToday.Text
        SetText vasID, gPatGen.Age, asRow, colPAge
    Else
        SetText vasID, 0, asRow, colPAge
        SetText vasID, 0, asRow, colState
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
End Function

Function Make_Order(argNo As String, argRow As Integer) As String
'Order Text 만들기
    Dim sRetOrder As String     'Order Text넣을 변수
    Dim sOrder As String
    
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim sExamCode As String     '검사코드
    Dim sRSCode As String       '검사항목코드
    Dim sEquipCode As String
    
    Dim sDate As String
    
    Dim iCnt_Ord As Integer    'Order conut
    Dim sReceNo As String
        
    Dim llRow As Long
    
    Dim sHead As String
    Dim sPatient As String

    sRetOrder = ""
    
    If argNo = "" Then
        Exit Function
    End If

    sDate = SeperatorCls(txtToday.Text)
    
    ClearSpread vasCode
    
    '검사코드, 검사항목코드 가져오기
    SQL = " Select DR_CODE From DEPARTDAT" & vbCrLf & _
          " Where DR_DATE = '" & Trim(GetText(vasID, argRow, colReqDate)) & "' " & vbCrLf & _
          " And DR_CHART = '" & Trim(argNo) & "'" & vbCrLf & _
          " And DR_CODE in (" & gAllExam & ") "
    res = db_select_Vas(gServer, SQL, vasCode)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    If res = 0 Then
        SetText vasID, "0", argRow, colOCnt   'Order 없음
        SetForeColor vasID, argRow, argRow, 255, 0, 0
    End If
    
    For iRow = 1 To vasCode.DataRowCnt
        SQL = " Select GD_CODE From GUMSADAT " & CR & _
              " Where GD_DATE = '" & Trim(GetText(vasID, argRow, colReqDate)) & "' " & vbCrLf & _
              " And GD_CHART = '" & Trim(argNo) & "'" & vbCrLf & _
              " And GD_CODE = '" & Trim(GetText(vasCode, iRow, 1)) & "' "
        res = db_select_Col(gServer, SQL)
        
        If gReadBuf(0) = Trim(GetText(vasCode, iRow, 1)) Then
            DeleteRow vasCode, iRow, iRow
            
            iRow = 1
        End If
    Next iRow
    
    'Order
    sOrder = ""
    iCnt_Ord = 0
    i = 1
    Do While i <= vasCode.DataRowCnt
    'For i = 1 To vasCode.DataRowCnt
        sEquipCode = ""
        sExamCode = Trim(GetText(vasCode, i, 1))
        sEquipCode = GetEquip_ExamCode(sExamCode)
        If sEquipCode <> "" Then
            'sEquipCode = Left(sEquipCode, 2)

            j = ScanCol(vasCode, sEquipCode, 2, 1)
            If j = -1 Then
                SetText vasCode, sEquipCode, i, 2
                iCnt_Ord = iCnt_Ord + 1
                sOrder = sOrder & sEquipCode

                i = i + 1
            ElseIf j > 0 Then
                DeleteRow vasCode, i, i
            End If
        End If
    'Next i
    Loop
    
    Make_Order = sOrder
    
    'SetText vasID, CStr(iCnt_Ord - 1), argRow, colOCnt
    SetText vasID, CStr(iCnt_Ord), argRow, colOCnt
    
End Function

Function SetResult(asResult As String, aiItem As Integer) As String
'DB에서 불러오기
    Dim iFloat As Integer

    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
    Case 7, 16
        iFloat = 2
    Case 14
        iFloat = 0
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
    End If

End Function

'''Function SetResult(asResult As String, asEquipCode As String)
'''    Dim i As Integer
'''    Dim sLVal As String
'''    Dim sHVal As String
'''    Dim sEquipCode As String
'''    Dim sEquipRes As String
'''    Dim sResult As String
'''    Dim sPoint As Integer
'''    Dim sResType As String
'''    Dim sResFlag As String
'''
'''
'''    sEquipRes = Trim(asResult)
'''    sEquipCode = Trim(asEquipCode)
'''    sResFlag = ""
'''
'''    If sEquipCode = "" Then
'''        Exit Function
'''    End If
'''
''''    If IsNumeric(sEquipRes) = False Then
''''        Exit Function
''''    End If
'''
'''    SQL = "select resprec, reflow, refhigh from EQPMASTER where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
'''    res = GetDBSelectColumn(gLocal, SQL)
'''
'''    If IsNumeric(gReadBuf(0)) = True Then
'''        sPoint = CInt(gReadBuf(0))
'''        sResType = ""
'''        For i = 0 To sPoint
'''            If i = 0 Then
'''                sResType = "#0"
'''            ElseIf i = 1 Then
'''                sResType = sResType & ".0"
'''            Else
'''                sResType = sResType & "0"
'''            End If
'''        Next
'''
'''        sResult = Format(sEquipRes, sResType)
'''    Else
'''        sResult = sEquipRes
'''    End If
'''
'''''    If IsNumeric(gReadBuf(1)) = True Then
'''''        sLVal = gReadBuf(1)
'''''        If CCur(sLVal) > CCur(sEquipRes) Then
'''''            sResFlag = "H"
'''''        End If
'''''    End If
'''''
'''''    If IsNumeric(gReadBuf(2)) = True Then
'''''        sHVal = gReadBuf(2)
'''''        If CCur(sHVal) < CCur(sEquipRes) Then
'''''            sResFlag = ">"
'''''        End If
'''''    End If
'''
'''    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
'''        sLVal = gReadBuf(1)
'''        sHVal = gReadBuf(2)
'''        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
'''            sResFlag = ""
'''        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
'''            sResFlag = "H"
'''        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
'''            sResFlag = "L"
'''        End If
'''    End If
'''
'''    gsFlag = sResFlag
'''    SetResult = sResult
'''
'''End Function

Private Sub MSComm1_OnComm()

    Dim s As String
    Dim sSendData As String
    
    s = MSComm1.Input
    
    Select Case s
    
        Case chrENQ     'Chr(5)
                
            Save_Raw_Data "[Rx:혈액학]" & s
            txtBuff2 = ""
            
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx:혈액학]" & chrACK
            
            txtBuff2.Text = ""
                   
        
        Case chrEOT
            
            Save_Raw_Data "[RX:혈액학]" & txtBuff2.Text & chrEOT
            
            'BC2800 txtBuff2
            
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx:혈액학]" & chrACK
        
        Case chrETX
            
            Save_Raw_Data "[RX:혈액학]" & txtBuff2.Text & chrETX
                
            'BC2800 txtBuff2
            
            MSComm1.Output = chrACK
            Save_Raw_Data "[Tx:혈액학]" & chrACK
            
            
        Case Else
            txtBuff2 = txtBuff2 & s
        
    End Select

End Sub

'-- BT 2000
Private Sub MSComm2_OnComm()
    
    Dim s As String
    Dim sSendData As String
    
    s = MSComm2.Input
        
    Select Case s
    Case chrSTX     'Chr(2)
        If gRecodeType = "Q" And Len(txtBuff) = 1 Then
            Save_Raw_Data "[Rx:생화학]" & txtBuff & s
            txtBuff = ""
            
            MSComm2.Output = chrSTX
            Save_Raw_Data "[Tx:생화학]" & chrSTX
            
            Exit Sub
        End If
        
        Save_Raw_Data "[RX:생화학]" & txtBuff & chrSTX
        
        MSComm2.Output = chrACK
        Save_Raw_Data "[TX:생화학]" & chrACK
        
        txtBuff.Text = ""
                
    Case chrEOT
        If gRecodeType = "Q" And Len(txtBuff) = 1 Then
            Save_Raw_Data "[Rx:생화학]" & txtBuff & s
            txtBuff = ""
            
            MSComm2.Output = chrSTX
            Save_Raw_Data "[Tx:생화학]" & chrSTX
                        
            Exit Sub
        End If
        
        Save_Raw_Data "[RX:생화학]" & txtBuff.Text & chrEOT
        
        BT2000 txtBuff
        
        MSComm2.Output = chrSTX
        Save_Raw_Data "[Tx:생화학]" & chrSTX
        
    Case chrACK
        If gRecodeType = "Q" And Len(txtBuff) = 1 Then
            Save_Raw_Data "[Rx:생화학]" & txtBuff & s
            txtBuff = ""
            
            MSComm2.Output = chrSTX
            Save_Raw_Data "[Tx:생화학]" & chrSTX
            
            
            Exit Sub
        End If
        
        Save_Raw_Data "[Rx]" & chrACK
        
        If gRecodeType = "Q" Then
            'gOrdRow = 0
            
            gOrdRow = gOrdRow + 1
    
            If gOrdRow <= vasOrder.DataRowCnt Then
    
                sSendData = GetText(vasOrder, gOrdRow, 1)
                
                MSComm2.Output = sSendData
                Save_Raw_Data "[Tx:생화학]" & sSendData
            Else
                ClearSpread vasOrder
                gOrdRow = 0
                gRecodeType = ""
                
                Me.MousePointer = 0
                
                MSComm2.Output = chrEOT
                Save_Raw_Data "[Tx:생화학]" & chrEOT
                
                Unload frmPatSear
            End If
        ElseIf gRecodeType = "R" Then
            MSComm2.Output = "R" & chrEOT
            Save_Raw_Data "[TX:생화학]" & "R" & chrEOT
        End If
        
        txtBuff.Text = ""
        
    Case Else
        txtBuff = txtBuff & s
        
        If gRecodeType = "Q" And Len(txtBuff) = 2 Then
            Save_Raw_Data "[Rx:생화학]" & txtBuff
            txtBuff = ""
            
            MSComm2.Output = chrSTX
            Save_Raw_Data "[Tx:생화학]" & chrSTX
                        
            Exit Sub
        End If
    End Select
    
End Sub

Private Sub sspMode_Click()
    If sspMode.Caption = "수정모드" Then
        sspMode.Caption = "전송모드"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "전송모드" Then
        sspMode.Caption = "수정모드"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If

End Sub

Private Sub subDel_Click()
    Dim i As Long
    Dim sSpecID As String
    Dim j As Integer
    
    i = vasID.ActiveRow
    
    sSpecID = Trim(GetText(vasID, i, colPID))
    
    vasID.DeleteRows i, 1
    If i > vasID.DataRowCnt Then
        i = vasID.DataRowCnt
    End If
    
    'vasID.MaxRows = vasID.DataRowCnt
    vasID.MaxRows = vasID.DataRowCnt + 1
    
    vasActiveCell vasID, i, colPID
    vasID.SetFocus
    
    '2004/06/11 이상은
    '검체번호 삭제함과 동시에 오더부분도 삭제
    For j = 1 To vasOrderBuf.DataRowCnt
        If sSpecID = Trim(GetText(vasOrderBuf, j, colPID)) Then
            vasOrderBuf.DeleteRows j, 1
            
            j = 1
        End If
        
        If j = vasOrderBuf.DataRowCnt Then
            Exit For
        End If
    Next j
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtBuff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BT2000 txtBuff
        txtBuff = ""
    End If
End Sub

Private Sub txtBuff2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BC2800 txtBuff2
        txtBuff2 = ""
    End If
End Sub

Private Sub txtMain_GotFocus(Index As Integer)

    txtMain(Index).BackColor = &HFFFFC0
    
    txtMain(Index).SelStart = 0
    txtMain(Index).SelLength = Len(txtMain(Index).Text)

End Sub

Private Sub txtMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If Index = 9 Then
                txtMain(0).SetFocus
            Else
                txtMain(Index + 1).SetFocus
            End If
            
        Case vbKeyEscape
            SSPanel2.Visible = False
    End Select

End Sub

Private Sub txtMain_LostFocus(Index As Integer)
    
    txtMain(Index).BackColor = &H80000005

End Sub

Private Sub txtMain1_GotFocus(Index As Integer)

    txtMain1(Index).BackColor = &HFFFFC0
    
    txtMain1(Index).SelStart = 0
    txtMain1(Index).SelLength = Len(txtMain(Index).Text)

End Sub

Private Sub txtMain1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If Index = 2 Then
                txtMain1(0).SetFocus
            Else
                txtMain1(Index + 1).SetFocus
            End If
            
        Case vbKeyEscape
            SSPanel2.Visible = False
    End Select

End Sub

Private Sub txtMain1_LostFocus(Index As Integer)

    txtMain1(Index).BackColor = &H80000005

End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
    
    lsID = Trim(GetText(vasID, Row, colPID))
    ClearSpread vasRes
    
'    SQL = "select '', barcode, equipcode,  examcode, examname, result, refflag, panicflag, deltaflag, unit, refvalue, panicvalue, result " & vbCrLf & _
'          "FROM pat_res " & vbCrLf & _
'          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'          "  AND Barcode = '" & Trim(GetText(vasID, vasID.Row, colPID)) & "' " & vbCrLf & _
'          "  order by equipcode"

'BT2000 과 BC2800동시에 출력하기 위해서...
    
    SQL = "select '', barcode, equipcode,  examcode, examname, result, refflag, panicflag, deltaflag, unit, refvalue, panicvalue, result " & vbCrLf & _
        "FROM pat_res " & vbCrLf & _
        "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
        "  AND Barcode = '" & Trim(GetText(vasID, vasID.Row, colPID)) & "' " & vbCrLf
    '-- 2012.06.26 조회조건 추가
    SQL = SQL & "  AND (result <> '' or result is null)    "
    SQL = SQL & "  order by equipcode"
      
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For i = 1 To vasRes.DataRowCnt
        '참조치
        Select Case Trim(GetText(vasRes, i, colRCheck))
        Case "H"
            SetText vasRes, "▲", i, 7
            
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            SetText vasRes, "▼", i, 7
                        
            vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 7
            vasRes.ForeColor = RGB(255, 255, 255)
        End Select
        
'        'Panic
'        Select Case Trim(GetText(vasRes, i, 8))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'
'        'Delta
'        Select Case Trim(GetText(vasRes, i, 9))
'        Case "D"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
    Next i

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim lsEquip As String
    
    sExamDate = GetDateFull
    
    lsEquip = asSend
    
'    If Trim(asSend) = "A" Then
'        gEquip = "BT2000"
'    Else
'        gEquip = "BC2800"
'    End If
    
    sCnt = ""
    'SQL = "delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & lsEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colPID)) & "' "
    
    'SaveQuery SQL
    'res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If
    
    
    
''          SQL = "select barcode "
''    SQL = SQL & "  from pat_res " & _
''                " WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
''                "   AND equipno = '" & lsEquip & "' " & vbCrLf & _
''                "   AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
''                "   AND barcode = '" & Trim(GetText(vasID, asRow1, colPID)) & "' "
''
''    res = db_select_Col(gLocal, SQL)
''
''    If res = 0 Then
''        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, receno, pid, " & _
''              "pname, pjumin, page, psex, resdate, " & _
''              "equipcode, examcode, examtype, result, sendflag, examname, " & _
''              "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
''              "VALUES ('" & Format(CDate(txtToday.Text), "yyyymmdd") & "', '" & Trim(lsEquip) & "', " & _
''              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', '', " & _
''              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
''              "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " & _
''              "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
''              "'" & sExamDate & "', " & vbCrLf & _
''              "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', '', " & _
''              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
''              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " & _
''              "'" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', " & _
''              "'" & Trim(GetText(vasRes, asRow2, colRef)) & "', '" & Trim(GetText(vasRes, asRow2, colPanic)) & "') "
''        res = SendQuery(gLocal, SQL)
''    Else
        SQL = " Update pat_res Set " & CR & _
              " examdate =  '" & Format(CDate(txtToday.Text), "yyyymmdd") & "', " & CR & _
              " result =  '" & Trim(GetText(vasRes, asRow2, colResult)) & "' " & CR & _
              " WHERE equipno = '" & lsEquip & "' " & vbCrLf & _
              "   AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
              "   AND barcode = '" & Trim(GetText(vasID, asRow1, colPID)) & "' "
        
        SaveQuery SQL
        
        res = SendQuery(gLocal, SQL)
''    End If

              '" WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
              "   AND equipno = '" & lsEquip & "' " & vbCrLf & _
              "   AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
              "   AND barcode = '" & Trim(GetText(vasID, asRow1, colPID)) & "' "

    

    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_One_CBC(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim lsEquip As String
    
    lsEquip = asSend
    
    sExamDate = GetDateFull
    
'    If Trim(asSend) = "A" Then
'        gEquip = "BT2000"
'    Else
'        gEquip = "BC2800"
'    End If
    
    sCnt = ""
    SQL = "delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & lsEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colPID)) & "' "
    
    'SaveQuery SQL
    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, receno, pid, " & _
          "pname, pjumin, page, psex, resdate, " & _
          "equipcode, examcode, examtype, result, sendflag, examname, " & _
          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(txtToday.Text), "yyyymmdd") & "', '" & Trim(lsEquip) & "', " & _
          "'" & Trim(GetText(vasID, asRow1, colPID)) & "', '', " & _
          "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " & _
          "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
          "'" & sExamDate & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', '', " & _
          "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " & _
          "'" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', " & _
          "'" & Trim(GetText(vasRes, asRow2, colRef)) & "', '" & Trim(GetText(vasRes, asRow2, colPanic)) & "') "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Private Sub vasID_KeyPress(KeyAscii As Integer)
    Dim sSpecID As String
    Dim llRow As Long
    Dim iRow As Long

    If KeyAscii = 13 Then

        llRow = vasID.Row
        sSpecID = Trim(GetText(vasID, llRow, colPID))
        
'        For iRow = 1 To vasID.DataRowCnt
'            If iRow <> llRow Then
'
'                If sSpecID = Trim(GetText(vasID, iRow, colPID)) Then
'                    SetText vasID, "", llRow, colPID
'                    vasActiveCell vasID, vasID.MaxRows, colPID
'                    vasID.SetFocus
'                    Exit Sub
'                End If
'            End If
'        Next
        
        '샘플의 환자 정보 가져오기
        Get_Sample_Info llRow
    
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop
End Sub


Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasID.ActiveRow
        If vasResCol = colResult And _
           Trim(GetText(vasRes, vasResRow, colResult)) <> Trim(GetText(vasRes, vasResRow, colResult1)) Then
            
            Response = MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
            If Response = vbYes Then
                '판정, 델타, 패닉 수정
                Check_Result Trim(GetText(vasID, vasIDRow, colPID)), _
                             Trim(GetText(vasID, vasIDRow, colPID)), _
                             Trim(GetText(vasRes, vasResRow, colExamCode)), _
                             Trim(GetText(vasRes, vasResRow, colResult)), _
                             vasResRow, Trim(GetText(vasID, vasIDRow, colPSex))

                SQL = " Update pat_res " & vbCrLf & _
                      " Set result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "', " & vbCrLf & _
                      " refFlag = '" & Trim(GetText(vasRes, vasResRow, colRCheck)) & "', " & vbCrLf & _
                      " panicFlag = '" & Trim(GetText(vasRes, vasResRow, colPCheck)) & "', " & vbCrLf & _
                      " deltaFlag = '" & Trim(GetText(vasRes, vasResRow, colDCheck)) & "' " & vbCrLf & _
                      " WHERE examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipExam)) & "'" & vbCrLf & _
                      "  AND barcode = '" & Trim(GetText(vasID, vasIDRow, colPID)) & "' "
                res = SendQuery(gLocal, SQL)
                
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                
            End If
        End If
        
    End If
End Sub

Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Select ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
          "From ExamMaster " & CR & _
          " Where HID = '115' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    sResClassCode = Trim(gReadBuf(0))
    
    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           Check_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(13)) Then
            If CInt(gReadBuf(13)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(13))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(1))
            sHigh = Trim(gReadBuf(2))
        Case "F"
            sLow = Trim(gReadBuf(3))
            sHigh = Trim(gReadBuf(4))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


'Panic 체크
        sPanicLow = ""
        sPanicHigh = ""
        
        sPanicGubun = Trim(gReadBuf(5))
        
        Select Case asSex
        Case "M", ""
            sPanicLow = Trim(gReadBuf(6))
            sPanicHigh = Trim(gReadBuf(7))
        Case "F"
            sPanicLow = Trim(gReadBuf(8))
            sPanicHigh = Trim(gReadBuf(9))
        End Select
        
        If sPanicGubun = "0" Then '상한/하한
            If sPanicLow = "" Or sPanicHigh = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) > CCur(sDiffRet) Then
                    PanicRet = "L"
                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
                    PanicRet = "H"
                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
                    PanicRet = ""
                End If
            End If
        ElseIf sPanicGubun = "1" Then 'percent
            If sPanicLow = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "L"
                    Else
                        PanicRet = ""
                    End If
                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "H"
                    Else
                        PanicRet = ""
                    End If
                Else
                    PanicRet = ""
                End If
            End If
        End If
        

'Delta 체크
        sDeltaLow = ""
        sDeltaHigh = ""
                
        sTmpRece1 = ""
        sTmpRet1 = ""
        sTmpRece2 = ""
        sTmpRet2 = ""
        PreResult = ""
        
        sMax_ReceNo = ""
'        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
        sReceNo = argBarCode
       
'2004/06/09 이상은
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
              
'2004/12/30 이상은 - 정렬부분 추가
        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
              " Where HID = '115' " & CR & _
              " And PID = '" & Trim(argPID) & "' " & CR & _
              " And ReceNo < '" & argBarCode & "' " & CR & _
              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
              " Group By Result " & CR & _
              " Order by 2 desc "
        res = db_select_Col(gServer, SQL)
              
        If res > 0 And gReadBuf(0) <> "" Then
            PreResult = gReadBuf(0)
        Else
            PreResult = ""
        End If
      
        If PreResult <> "" Then
          'PreResult = Trim(gReadBuf(0))
          sDeltaGubun = Trim(gReadBuf(10))
          
          sDeltaLow = Trim(gReadBuf(11))
          sDeltaHigh = Trim(gReadBuf(12))
          
            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
'            sDiffRet = PreResult - sDiffRet
            sDiffRet1 = sDiffRet - PreResult
            If sDeltaGubun = "0" Then '상한/하한
                If sDeltaLow = "" Or sDeltaHigh = "" Then
                    DeltaRet = ""
                Else
                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
                        DeltaRet = "L"
                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
                        DeltaRet = "H"
                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
                        DeltaRet = ""
                    End If
                End If
              
            ElseIf sDeltaGubun = "1" Then 'percent
               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
                  DeltaRet = ""
               Else
                   If sDeltaLow = "" Then
                        DeltaRet = ""
                    Else
                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
                            DeltaRet = "D"
                        Else
                            DeltaRet = ""
                        End If
                    End If
               End If
            End If
        End If
        
    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "양성"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    SetText vasRes, PanicRet, argRow, colPCheck
    SetText vasRes, DeltaRet, argRow, colDCheck
    

    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If PanicRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If DeltaRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    ElseIf DeltaRet = "D" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    Check_Result = 1

End Function

'양식출력
''Public Sub Print_Report()
''
''    Dim i As Integer
''    Dim j As Integer
''    Dim iLeft As Integer
''    Dim iRow As Integer
''    Dim iRRow As Integer
''    Dim iCCow As Integer
''    Dim iRight As Integer
''    Dim iX As Integer
''    Dim iY As Integer
''    Dim iTmp As Integer
''
''    Dim iTmp1 As Integer
''    Dim iTmp2 As Integer
''    Dim iTmP5 As Integer
''
''    Dim sTmpResVal As String
''    Dim iTmpPos As Integer
''    Dim iTmpPos1 As Integer
''    Dim sTmpNam As String
''    Dim iRowCnt As Integer
''    Dim bTmp As Boolean
''
''    Dim sTmpPatNam As String
''    Dim lTmpChtNum As String
''    Dim sTmpLabDat As String
''    Dim sTmpSexTyp As String
''    Dim sTmpHigLow As String
''
''    Dim sTmpLowVal As String
''    Dim sTmpHigVal As String
''    'Dim sTmpAddVal(10) As String
''    Dim sTmpAge As String
''
''    Dim sSql As String
''
''    ' 기본데이타 올림
''    Printer.Font = "바탕체"
''
''    Printer.FontSize = 10
''
''    ' 양식부터 그린다.
''    iLeft = 300
''    iRight = 10100
''    iRow = 600
''    iRRow = 620
''
''    iX = iLeft
''    iY = 200
''
''    '가로줄
''    For i = 1 To 93
''        iCCow = 100
''        iCCow = iCCow * i
''        Call sSetPrint1(iX + iCCow, iY, "_")
''    Next
''
''    For i = 1 To 93
''        iCCow = 100
''        iCCow = iCCow * i
''        Call sSetPrint1(iX + iCCow, iY + 400, "_")
''    Next
''
''    iTmp1 = iY + 400
''
''    For j = 1 To 3
''        iRRow = 400
''        iRRow = (iRRow * j) + 400
''
''        For i = 1 To 93
''            iCCow = 100
''            iCCow = iCCow * i
''            Call sSetPrint1(iX + iCCow, iY + iRRow, "_")
''        Next
''
''        If j = 1 Then
''            iTmp2 = iY + iRRow
''        End If
''    Next
''
''    iTmp = iX + iCCow
''    '세로줄
''    For i = 1 To 79
''        iRRow = 170
''        iRRow = iRRow * i
''        Call sSetPrint1(iX + 50, iY + iRRow, "|")
''    Next
''
''    For i = 1 To 79
''        iRRow = 170
''        iRRow = iRRow * i
''        Call sSetPrint1(iTmp + 50, iY + iRRow, "|")
''    Next
''
''    iTmP5 = iY + iRRow    '마지막줄 변수저장
''
''    vasID.Row = vasID.ActiveRow
''    vasID.Col = 5
''    If Trim(vasID.Text) = "" Then
''        MsgBox "출력하실 내역을 선택하십시오", vbExclamation, "내역선택오류"
''        Exit Sub
''    Else
''        lTmpChtNum = vasID.Text
''    End If
''
''    vasID.Col = 6
''    sTmpPatNam = vasID.Text
''
''    vasID.Col = 8
''    sTmpSexTyp = vasID.Text
''
''    vasID.Col = 9
''    sTmpAge = vasID.Text
''
''    Printer.FontSize = 12
''    Printer.FontBold = True
''    Printer.FontName = "굴림체"
''    Call sSetPrint1(2800, 450, "이성우내과의원 임상병리 검사 결과지")
''
''    Printer.FontSize = 10
''    Printer.FontBold = False
''
''    Call sSetPrint1(600, 900, "수 신 자 명")
''    Call sSetPrint1(5250, 900, "챠 트 번 호")
''    Call sSetPrint1(600, 1300, "검사 의뢰일")
''    Call sSetPrint1(5500, 1300, "나   이")
''    Call sSetPrint1(7900, 1300, "성별")
''
''    Call sSetPrint1(600, 2100, "검사명" & "          " & "결과치" & "      " & "비교값")
''    Call sSetPrint1(600, 2230, "------------------------------------------")
''    Call sSetPrint1(5150, 2100, "검사명" & "                " & "결과치" & "      " & "비교값")
''    Call sSetPrint1(5150, 2230, "-------------------------------------------")
''
''    Call sSetPrint1(600, 11100, "검사명" & "         " & "결과치" & "      " & "비교값")
''    Call sSetPrint1(600, 11300, "-----------------------------------------")
''    Call sSetPrint1(5300, 11100, "검사명" & "                " & "결과치" & "   " & "비교값")
''    Call sSetPrint1(5300, 11300, "-----------------------------------------")
''
''    Call sSetPrint1(6800, 900, lTmpChtNum)
''    Call sSetPrint1(2200, 900, sTmpPatNam)
''    Call sSetPrint1(8900, 1300, sTmpSexTyp)
''    Call sSetPrint1(2200, 1300, txtToday.Text)
''    Call sSetPrint1(7100, 1300, sTmpAge)
''
''    Printer.FontSize = 13
''
''    For i = 1 To 7
''        iRRow = 92
''        iRRow = iRRow * i
''        Call sSetPrint1(1900, (iTmp1 + 90) + iRRow, "|")
''    Next
''
''    For i = 1 To 7
''        iRRow = 92
''        iRRow = iRRow * i
''        Call sSetPrint1(4950, (iTmp1 + 90) + iRRow, "|")
''    Next
''
''    For i = 1 To 7
''        iRRow = 92
''        iRRow = iRRow * i
''        Call sSetPrint1(6550, (iTmp1 + 90) + iRRow, "|")
''    Next
''
''    For i = 1 To 3
''        iRRow = 92
''        iRRow = iRRow * i
''        Call sSetPrint1(7700, (iTmp2 + 90) + iRRow, "|")
''    Next
''
''
''    For i = 1 To 3
''        iRRow = 80
''        iRRow = iRRow * i
''        Call sSetPrint1(8500, (iTmp2 + 90) + iRRow, "|")
''    Next
''
''
''    For i = 1 To 93
''        iCCow = 100
''        iCCow = iCCow * i
''        Call sSetPrint1(iX + iCCow, iTmP5 - 40, "_")
''    Next
''
''    Printer.Line (510, 2050)-(4950, 10900), , B
''    Printer.Line (5000, 2050)-(9600, 10900), , B
''    Printer.Line (510, 11000)-(9600, 13600), , B
''    Printer.Line (5000, 11100)-(5000, 13500)
''
''    Printer.FontSize = 10
''    Printer.FontBold = True
''
''    '아래 소변검사 출력
''    Call sSetPrint1(600, 11500, "Urobilinogen")
''    Call sSetPrint1(2300, 11500, txtMain(TXT_URO).Text & "           음성")
''
''    Call sSetPrint1(600, 11900, "Bilirubin")
''    Call sSetPrint1(2300, 11900, txtMain(TXT_BIL).Text & "           음성")
''
''    Call sSetPrint1(600, 12300, "Ketone")
''    Call sSetPrint1(2300, 12300, txtMain(TXT_KET).Text & "           음성")
''
''    Call sSetPrint1(600, 12700, "Blood")
''    Call sSetPrint1(2300, 12700, txtMain(TXT_RBC).Text & "           음성")
''
''    Call sSetPrint1(600, 13100, "Protein")
''    Call sSetPrint1(2300, 13100, txtMain(TXT_PRO).Text & "           음성")
''
''    Call sSetPrint1(5300, 11500, "Nitrite")
''    Call sSetPrint1(7600, 11500, txtMain(TXT_NIT).Text & "        음성")
''
''    Call sSetPrint1(5300, 11900, "Leukocytes")
''    Call sSetPrint1(7600, 11900, txtMain(TXT_LEU).Text & "        음성")
''
''    Call sSetPrint1(5300, 12300, "Glucose")
''    Call sSetPrint1(7600, 12300, txtMain(TXT_GLU).Text & "        음성")
''
''    Call sSetPrint1(5300, 12700, "Specific Gravity")
''    Call sSetPrint1(7600, 12700, txtMain(TXT_SPE).Text & "")
''
''    Call sSetPrint1(5300, 13100, "pH")
''    Call sSetPrint1(7600, 13100, txtMain(TXT_PH1).Text & "        5~8")
''    '붉은판넬
''    Call sSetPrint1(2300, 9200, txtMain1(TXT_HBA).Text)
''    Call sSetPrint1(2300, 10100, txtMain1(TXT_HBG).Text)
''    Call sSetPrint1(2300, 10500, txtMain1(TXT_HBB).Text)
''
''           sSql = "Select * From equipexam "
'''    sSql = sSql & " Where equipno  = 'BT2000' "
''    sSql = sSql & " Order by seqno Asc "
''
''    Set cmdSQL.ActiveConnection = cn
''    cmdSQL.CommandText = sSql
''    Set RS = cmdSQL.Execute
''
''    iTmpPos = 2000
''    iTmpPos1 = 2000
''    iRowCnt = 1
''
''    If Not (RS.BOF And RS.EOF) Then
''        Do Until RS.EOF
''            iTmpPos = iTmpPos + 450
''
''            For iRowCnt = 1 To vasRes.MaxRows
''                vasRes.Row = iRowCnt
''                vasRes.Col = 3
''
''                If Trim(RS!equipcode) = Trim(vasRes.Text) Then
''                    bTmp = True
''                End If
''
''                vasRes.Col = 6
''                If Trim(vasRes.Text) = "" Then
''                Else
''                    sTmpResVal = vasRes.Text
''                End If
''
''                If bTmp = True Then
''                    Exit For    '결과가 있으면 빠져나간다.
''                End If
''            Next
''
''            sTmpNam = RS!examname
''
''            ' 비교치가 Null 일경우 0으로 셋팅한다.
''            If IsNull(RS!refhigh) = True And IsNull(RS!refhigh) = False Then
''                sTmpHigVal = 0
''            ElseIf IsNull(RS!reflow) = True And IsNull(RS!reflow) = False Then
''                sTmpLowVal = 0
''            ElseIf IsNull(RS!reflow) = True And IsNull(RS!reflow) = True Then
''                sTmpHigVal = 0
''                sTmpLowVal = 0
''            Else
''                sTmpHigVal = RS!refhigh
''                sTmpLowVal = RS!reflow
''            End If
''
''            '검사가 없는항목은 표시하지 않는다.
''            If bTmp = True Then
''                sTmpHigLow = FindResVal(sTmpResVal, sTmpLowVal, sTmpHigVal)
''            Else
''                sTmpHigLow = ""
''            End If
''
''            If iTmpPos >= 13500 Or RS!seqno >= 20000 Then
''
''                iTmpPos1 = iTmpPos1 + 450
''
''                If bTmp = True Then
''                    Call sSetPrint1(5150, iTmpPos1, sTmpNam)
''                    Call sSetPrint1(7350, iTmpPos1, sTmpResVal)
''                    Call sSetPrint1(8450, iTmpPos1, sTmpLowVal & "-" & sTmpHigVal)
''                    Call sSetPrint1(8050, iTmpPos1, sTmpHigLow)
''                Else
''                    Call sSetPrint1(5150, iTmpPos1, sTmpNam)
''                    Call sSetPrint1(7350, iTmpPos1, "  ")
''                    Call sSetPrint1(8450, iTmpPos1, sTmpLowVal & "-" & sTmpHigVal)
''                    Call sSetPrint1(8050, iTmpPos1, sTmpHigLow)
''                End If
''            Else
''                If bTmp = True Then
''                    Call sSetPrint1(600, iTmpPos, sTmpNam)
''                    Call sSetPrint1(2450, iTmpPos, sTmpResVal)
''                    Call sSetPrint1(3400, iTmpPos, sTmpLowVal & "-" & sTmpHigVal)
''                    Call sSetPrint1(2950, iTmpPos, sTmpHigLow)
''                Else
''                    Call sSetPrint1(600, iTmpPos, sTmpNam)
''                    Call sSetPrint1(2450, iTmpPos, "  ")
''                    Call sSetPrint1(3400, iTmpPos, sTmpLowVal & "-" & sTmpHigVal)
''                    Call sSetPrint1(2950, iTmpPos, sTmpHigLow)
''                End If
''            End If
''
''            bTmp = False
''
''        RS.MoveNext
''        Loop
''
''        sTmpNam = ""
''    End If
''
''    Printer.EndDoc
''
''End Sub

' 설정된 좌표에 데이타 출력
Public Sub sSetPrint1(iPrmX As Integer, iPrmY As Integer, sPrmData As String)
    
    ' Y좌표, Tab값, 출력값
    Dim iLength As Integer
    
    Printer.CurrentX = iPrmX
    Printer.CurrentY = iPrmY
    
    Printer.Print sPrmData
    
End Sub

Function FindResVal(ByVal iPrmResVal As Integer, ByVal iPrmLowVal As Integer, ByVal iPrmHigVal As Integer) As String
    
    If iPrmResVal = 0 Then
        FindResVal = ""
    ElseIf iPrmLowVal > iPrmResVal Then
        FindResVal = "L"
    ElseIf iPrmHigVal < iPrmResVal Then
        FindResVal = "H"
    Else
        FindResVal = ""
    End If
          
End Function
