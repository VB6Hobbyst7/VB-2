VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00F8E4D8&
   Caption         =   "OK SOFT"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   15960
   StartUpPosition =   1  '소유자 가운데
   WindowState     =   2  '최대화
   Begin VB.Frame FraHidden 
      Caption         =   "HIDDEN CONTROL"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8805
      Left            =   11460
      TabIndex        =   84
      Top             =   1980
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Frame fraLinks 
         Caption         =   " Result File List "
         Height          =   6000
         Left            =   12030
         TabIndex        =   203
         Top             =   2580
         Width           =   10965
         Begin VB.TextBox txtLink 
            Height          =   5565
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   206
            Top             =   240
            Width           =   10635
         End
         Begin VB.TextBox txtHTML 
            Height          =   2325
            Left            =   5520
            MultiLine       =   -1  'True
            TabIndex        =   205
            Top             =   3030
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   2535
            Left            =   5550
            MultiLine       =   -1  'True
            TabIndex        =   204
            Top             =   420
            Visible         =   0   'False
            Width           =   4755
         End
         Begin VB.Timer tmrRefresh 
            Left            =   5940
            Top             =   150
         End
      End
      Begin VB.Frame fraSetup 
         Caption         =   "Setup"
         Height          =   1305
         Left            =   11430
         TabIndex        =   191
         Top             =   510
         Visible         =   0   'False
         Width           =   11055
         Begin VB.CommandButton cmdReadJSON 
            Caption         =   "Read JSON Data From File"
            Height          =   435
            Left            =   3840
            TabIndex        =   201
            Top             =   270
            Width           =   3705
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "Open"
            Height          =   375
            Left            =   1860
            TabIndex        =   200
            Top             =   270
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdParse 
            Caption         =   "Parse"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2865
            TabIndex        =   199
            Top             =   270
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CheckBox chkFullLink 
            Caption         =   "Full Link"
            Height          =   285
            Left            =   180
            TabIndex        =   198
            Top             =   900
            Width           =   1275
         End
         Begin VB.CheckBox chkLink 
            Caption         =   "Link"
            Height          =   285
            Left            =   1950
            TabIndex        =   197
            Top             =   900
            Value           =   1  '확인
            Width           =   960
         End
         Begin VB.CheckBox chkImbed 
            Caption         =   "Imbed"
            Height          =   285
            Left            =   3870
            TabIndex        =   196
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkPlugin 
            Caption         =   "PlugIn"
            Height          =   285
            Left            =   4830
            TabIndex        =   195
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkImage 
            Caption         =   "Image"
            Height          =   285
            Left            =   6750
            TabIndex        =   194
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkScript 
            Caption         =   "Script"
            Height          =   285
            Left            =   7770
            TabIndex        =   193
            Top             =   900
            Width           =   885
         End
         Begin VB.CheckBox chkForms 
            Caption         =   "Form"
            Height          =   285
            Left            =   5820
            TabIndex        =   192
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label4 
            Caption         =   "Page URL :"
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdTVSave 
         Caption         =   "저장"
         Height          =   345
         Left            =   9990
         TabIndex        =   188
         Top             =   4560
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtTV 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8520
         TabIndex        =   187
         Top             =   4590
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdView 
         Appearance      =   0  '평면
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10230
         TabIndex        =   180
         Top             =   6810
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtURL 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5580
         TabIndex        =   179
         Text            =   "C:\프로젝트\엘리어i\Alerei Test Results.htm"
         Top             =   6840
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.CommandButton cmdResult 
         BackColor       =   &H00C0FFFF&
         Caption         =   "결과얻기"
         Height          =   405
         Left            =   7410
         Style           =   1  '그래픽
         TabIndex        =   178
         Top             =   2250
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtTimer 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6510
         TabIndex        =   177
         Text            =   "60"
         Top             =   2250
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Timer tmrResult 
         Left            =   6870
         Top             =   450
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7110
         TabIndex        =   170
         Top             =   1710
         Width           =   2985
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6030
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Timer tmrFlipFlop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3120
         Top             =   720
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "종료"
         Height          =   315
         Left            =   6750
         TabIndex        =   168
         Top             =   1140
         Width           =   795
      End
      Begin VB.Timer tmrComm 
         Enabled         =   0   'False
         Left            =   2670
         Top             =   720
      End
      Begin VB.Frame frameCutOff 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   5070
         TabIndex        =   157
         Top             =   2850
         Visible         =   0   'False
         Width           =   5175
         Begin VB.ComboBox cboCOL 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":0E42
            Left            =   2730
            List            =   "frmMain.frx":0E44
            TabIndex        =   164
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtCOLIn 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            TabIndex        =   163
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox txtCOLOut 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            TabIndex        =   162
            Top             =   300
            Width           =   1545
         End
         Begin VB.TextBox txtCOMOut 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            TabIndex        =   161
            Top             =   660
            Width           =   1545
         End
         Begin VB.ComboBox cboCOH 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":0E46
            Left            =   2730
            List            =   "frmMain.frx":0E48
            TabIndex        =   160
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtCOHIn 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            TabIndex        =   159
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtCOHOut 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            TabIndex        =   158
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   12
            Left            =   210
            Picture         =   "frmMain.frx":0E4A
            Top             =   360
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (L)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   20
            Left            =   480
            TabIndex        =   167
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (M)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   480
            TabIndex        =   166
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff (H)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   480
            TabIndex        =   165
            Top             =   1110
            Width           =   840
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   9
            Left            =   210
            Picture         =   "frmMain.frx":1234
            Top             =   720
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   13
            Left            =   210
            Picture         =   "frmMain.frx":161E
            Top             =   1080
            Width           =   150
         End
      End
      Begin VB.Timer TimerVESCUBE 
         Left            =   780
         Top             =   360
      End
      Begin VB.Timer Timer2120 
         Enabled         =   0   'False
         Left            =   360
         Top             =   330
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   330
         TabIndex        =   127
         Top             =   6480
         Visible         =   0   'False
         Width           =   5445
         Begin VB.TextBox txtInstrument 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1230
            TabIndex        =   137
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtLab 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1230
            TabIndex        =   136
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1230
            TabIndex        =   135
            Top             =   1380
            Width           =   1185
         End
         Begin VB.TextBox txtReagent 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3690
            TabIndex        =   134
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtMethod 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3690
            TabIndex        =   133
            Top             =   660
            Width           =   1185
         End
         Begin VB.TextBox txtLot 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3690
            TabIndex        =   132
            Top             =   300
            Width           =   1185
         End
         Begin VB.TextBox txtTemp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3690
            TabIndex        =   131
            Top             =   1380
            Width           =   1185
         End
         Begin VB.CommandButton cmdLabFind 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2430
            TabIndex        =   130
            Top             =   300
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton Command3 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5040
            TabIndex        =   129
            Top             =   300
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAnalyteFind 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2430
            TabIndex        =   128
            Top             =   630
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   17
            Left            =   210
            Picture         =   "frmMain.frx":1A08
            Top             =   1080
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "기기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   144
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Lab"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   480
            TabIndex        =   143
            Top             =   390
            Width           =   315
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   25
            Left            =   210
            Picture         =   "frmMain.frx":1DF2
            Top             =   360
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   26
            Left            =   210
            Picture         =   "frmMain.frx":21DC
            Top             =   1440
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "단위"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   480
            TabIndex        =   142
            Top             =   1470
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   27
            Left            =   2670
            Picture         =   "frmMain.frx":25C6
            Top             =   1080
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   28
            Left            =   2670
            Picture         =   "frmMain.frx":29B0
            Top             =   720
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "시약"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   36
            Left            =   2940
            TabIndex        =   141
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Method"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   37
            Left            =   2940
            TabIndex        =   140
            Top             =   750
            Width           =   630
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Lot"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   38
            Left            =   2940
            TabIndex        =   139
            Top             =   390
            Width           =   255
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   29
            Left            =   2670
            Picture         =   "frmMain.frx":2D9A
            Top             =   360
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   30
            Left            =   2670
            Picture         =   "frmMain.frx":3184
            Top             =   1440
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "온도"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   39
            Left            =   2940
            TabIndex        =   138
            Top             =   1470
            Width           =   360
         End
      End
      Begin VB.Frame frameSet 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 시스템 설정 "
         Height          =   1935
         Left            =   270
         TabIndex        =   117
         Top             =   4470
         Visible         =   0   'False
         Width           =   5025
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   119
            Text            =   "Combo1"
            Top             =   510
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   118
            Text            =   "Combo1"
            Top             =   1110
            Width           =   2295
         End
         Begin VB.Image Image1 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":356E
            Top             =   540
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   600
            TabIndex        =   123
            Top             =   570
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "프로토콜"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   600
            TabIndex        =   122
            Top             =   1170
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   225
            Left            =   390
            Picture         =   "frmMain.frx":3958
            Top             =   1140
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   4110
            TabIndex        =   121
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "OCS"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   4110
            TabIndex        =   120
            Top             =   1170
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "시스템설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3660
         TabIndex        =   115
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   103
         Top             =   1140
         Width           =   3045
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Seq 사용"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1770
            TabIndex        =   105
            Top             =   90
            Width           =   1155
         End
         Begin VB.OptionButton optBarSeq 
            BackColor       =   &H00FFFFFF&
            Caption         =   "검체번호 사용"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   104
            Top             =   90
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1470
         TabIndex        =   98
         Top             =   2040
         Width           =   2565
         Begin VB.OptionButton optSaveResult 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LIS결과"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1260
            TabIndex        =   100
            Top             =   30
            Width           =   1095
         End
         Begin VB.OptionButton optSaveResult 
            BackColor       =   &H00FFFFFF&
            Caption         =   "장비결과"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   90
            TabIndex        =   99
            Top             =   30
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   95
         Top             =   1620
         Width           =   1875
         Begin VB.OptionButton optTrans 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   90
            TabIndex        =   97
            Top             =   30
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optTrans 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수동"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   930
            TabIndex        =   96
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2100
         Top             =   300
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2580
         Top             =   300
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   1380
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3D42
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":42DC
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4876
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4E10
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":56A2
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":57FC
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5956
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   885
         Left            =   300
         TabIndex        =   114
         Top             =   2490
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   1561
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
         SpreadDesigner  =   "frmMain.frx":5AB0
      End
      Begin HAN2ENGOCXLib.Han2EngOCX Han2Eng 
         Height          =   315
         Left            =   3090
         TabIndex        =   146
         Top             =   360
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   0
      End
      Begin FPSpread.vaSpread spdQcResult 
         Height          =   825
         Left            =   300
         TabIndex        =   151
         Top             =   3450
         Visible         =   0   'False
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   1455
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
         SpreadDesigner  =   "frmMain.frx":5CF7
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   8550
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   7830
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   7800
         TabIndex        =   174
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   128516097
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   9360
         TabIndex        =   175
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   128516097
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "total volume(L)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   42
         Left            =   6990
         TabIndex        =   189
         Top             =   4665
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Image Image5 
         Height          =   225
         Index           =   31
         Left            =   6720
         Picture         =   "frmMain.frx":5F3E
         Top             =   4635
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblURL 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "Page URL :"
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
         Left            =   5610
         TabIndex        =   181
         Top             =   6510
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   43
         Left            =   9150
         TabIndex        =   176
         Top             =   1200
         Width           =   150
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   5640
         TabIndex        =   156
         Top             =   1410
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   5670
         TabIndex        =   155
         Top             =   1140
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   5640
         TabIndex        =   154
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   5670
         TabIndex        =   153
         Top             =   630
         Width           =   2580
      End
      Begin VB.Label lblPatInfo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "홍길"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   5640
         TabIndex        =   152
         Top             =   900
         Width           =   2820
      End
      Begin VB.Image imgDelete 
         Height          =   1260
         Left            =   6030
         Picture         =   "frmMain.frx":6328
         Top             =   7410
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Image imgSave 
         Height          =   1260
         Left            =   7440
         Picture         =   "frmMain.frx":8142
         Top             =   7530
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드사용"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   106
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   102
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과전송"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   101
         Top             =   1710
         Width           =   780
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source"
      Height          =   2700
      Left            =   5370
      TabIndex        =   171
      Top             =   9780
      Visible         =   0   'False
      Width           =   11025
      Begin VB.TextBox txtSource 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   172
         Top             =   480
         Width           =   10785
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title : "
         Height          =   255
         Left            =   240
         TabIndex        =   173
         Top             =   270
         Width           =   7695
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   15930
      TabIndex        =   2
      Top             =   1035
      Width           =   15960
      Begin VB.Frame fraResult 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   90
         Top             =   -60
         Visible         =   0   'False
         Width           =   14145
         Begin VB.ComboBox cboRstType 
            Appearance      =   0  '평면
            Height          =   300
            ItemData        =   "frmMain.frx":9E8B
            Left            =   420
            List            =   "frmMain.frx":9E8D
            TabIndex        =   110
            Top             =   180
            Width           =   1245
         End
         Begin VB.ComboBox cboState 
            Height          =   300
            ItemData        =   "frmMain.frx":9E8F
            Left            =   4710
            List            =   "frmMain.frx":9E91
            TabIndex        =   109
            Top             =   180
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   1770
            TabIndex        =   92
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   128516097
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   3330
            TabIndex        =   93
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   128516097
            CurrentDate     =   40457
         End
         Begin VB.Shape shpRS 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   7650
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblRSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "선택저장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7770
            TabIndex        =   190
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblRClear 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "화면정리"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9240
            TabIndex        =   147
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpRC 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   9120
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "~"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   3120
            TabIndex        =   94
            Top             =   240
            Width           =   150
         End
         Begin VB.Image imgGbn 
            Height          =   225
            Left            =   180
            Picture         =   "frmMain.frx":9E93
            Top             =   210
            Width           =   150
         End
         Begin VB.Shape shpR 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   6180
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과조회"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6300
            TabIndex        =   91
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame fraInterface 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   6510
         TabIndex        =   78
         Top             =   -60
         Width           =   14145
         Begin VB.TextBox txtBarcode 
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
            Height          =   375
            Left            =   9660
            TabIndex        =   169
            Text            =   "1234567890"
            Top             =   180
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton cmdInit 
            Caption         =   "초기화"
            Height          =   375
            Left            =   12210
            TabIndex        =   145
            Top             =   150
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Shape shpC 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   3300
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblClear 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "화면정리"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3480
            TabIndex        =   83
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpS 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   1860
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label lblSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "선택저장"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1980
            TabIndex        =   82
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblWork 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "워크조회"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   79
            Top             =   240
            Width           =   1125
         End
         Begin VB.Shape shpW 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   420
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "통신설정"
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
         Index           =   3
         Left            =   4830
         TabIndex        =   34
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   3
         Left            =   4710
         Top             =   60
         Width           =   1395
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   2
         Left            =   3240
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사설정"
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
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   27
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   1
         Left            =   1770
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과조회"
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
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   26
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape shpB 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00C0FFC0&
         Height          =   435
         Index           =   0
         Left            =   270
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "인터페이스"
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
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   25
         Top             =   150
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   15960
      TabIndex        =   0
      Top             =   0
      Width           =   15960
      Begin VB.Frame fraCommTest 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   15690
         TabIndex        =   207
         Top             =   60
         Visible         =   0   'False
         Width           =   4935
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   60
            TabIndex        =   209
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox txtRcv 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   450
            MultiLine       =   -1  'True
            TabIndex        =   208
            Top             =   120
            Width           =   4425
         End
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   10020
         TabIndex        =   107
         Top             =   450
         Width           =   2505
         _ExtentX        =   4419
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
         Format          =   128516096
         CurrentDate     =   40457
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   12720
         TabIndex        =   186
         Top             =   690
         Width           =   4275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14670
         TabIndex        =   185
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   13785
         TabIndex        =   184
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   12810
         TabIndex        =   183
         Top             =   330
         Width           =   360
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15210
         Picture         =   "frmMain.frx":A27D
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14295
         Picture         =   "frmMain.frx":A807
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   13350
         Picture         =   "frmMain.frx":AD91
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblCommStatus 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Com"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   16620
         TabIndex        =   182
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Index           =   27
         Left            =   9150
         TabIndex        =   108
         Top             =   540
         Width           =   720
      End
      Begin VB.Image Image7 
         Height          =   225
         Left            =   8880
         Picture         =   "frmMain.frx":B31B
         Top             =   510
         Width           =   150
      End
      Begin VB.Label lblHospInfo 
         BackStyle       =   0  '투명
         Caption         =   "전남대학교병원 HITACHI 7020[H36] 홍길동[12345]"
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
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   10485
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmMain.frx":B705
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.Frame frame2 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   780
      TabIndex        =   85
      Top             =   3570
      Visible         =   0   'False
      Width           =   20685
      Begin VB.CommandButton cmdRSL 
         Appearance      =   0  '평면
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   112
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkRAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   111
         Top             =   240
         Width           =   195
      End
      Begin FPSpread.vaSpread spdRResult 
         Height          =   9360
         Left            =   13620
         TabIndex        =   89
         Top             =   180
         Width           =   6960
         _Version        =   393216
         _ExtentX        =   12277
         _ExtentY        =   16510
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmMain.frx":CE48
         TextTip         =   2
      End
      Begin FPSpread.vaSpread spdROrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   88
         Top             =   180
         Width           =   13485
         _Version        =   393216
         _ExtentX        =   23786
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   20
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
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmMain.frx":D842
         UserResize      =   2
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  '평면
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   87
         Top             =   210
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   570
         TabIndex        =   86
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   50
      TabIndex        =   3
      Top             =   1650
      Width           =   20685
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   81
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton cmdSL 
         Appearance      =   0  '평면
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
      Begin FPSpread.vaSpread spdOrder 
         Height          =   9375
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   17235
         _Version        =   393216
         _ExtentX        =   30401
         _ExtentY        =   16536
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   20
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
         MaxCols         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmMain.frx":11DCB
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdResult 
         Height          =   9360
         Left            =   17370
         TabIndex        =   4
         Top             =   180
         Width           =   3210
         _Version        =   393216
         _ExtentX        =   5662
         _ExtentY        =   16510
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   50
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmMain.frx":1635C
         TextTip         =   2
      End
   End
   Begin VB.Frame frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   1230
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   20685
      Begin VB.Frame frameTestSet 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9315
         Left            =   14730
         TabIndex        =   8
         Top             =   180
         Width           =   5835
         Begin VB.CheckBox chkResSpec 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용유무"
            Height          =   390
            Left            =   4050
            TabIndex        =   150
            Top             =   3540
            Width           =   1185
         End
         Begin VB.CommandButton cmdQCMaster 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "Biorad QC 설정"
            Height          =   345
            Left            =   3870
            Style           =   1  '그래픽
            TabIndex        =   126
            Top             =   4830
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox txtAnalyte 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   124
            Top             =   4860
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Frame frameOrder 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   210
            TabIndex        =   75
            Top             =   6960
            Visible         =   0   'False
            Width           =   2085
            Begin VB.CommandButton cmdDelete 
               Appearance      =   0  '평면
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   420
               TabIndex        =   80
               Top             =   210
               Width           =   285
            End
            Begin VB.CommandButton cmdAppend 
               Appearance      =   0  '평면
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   76
               Top             =   210
               Width           =   285
            End
            Begin FPSpread.vaSpread spdOrdMst 
               Height          =   1920
               Left            =   90
               TabIndex        =   77
               Top             =   180
               Width           =   1890
               _Version        =   393216
               _ExtentX        =   3334
               _ExtentY        =   3387
               _StockProps     =   64
               BackColorStyle  =   1
               DisplayRowHeaders=   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   1
               MaxRows         =   50
               OperationMode   =   2
               RetainSelBlock  =   0   'False
               ScrollBars      =   2
               SelectBlockOptions=   0
               ShadowColor     =   13697023
               SpreadDesigner  =   "frmMain.frx":16DDC
               TextTip         =   2
            End
         End
         Begin VB.ComboBox cboResultType 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMain.frx":17353
            Left            =   1650
            List            =   "frmMain.frx":17355
            TabIndex        =   31
            Top             =   5220
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtRChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   24
            Top             =   1770
            Width           =   2115
         End
         Begin VB.TextBox txtEqpCD 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox txtTestCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   20
            Top             =   2220
            Width           =   2115
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   1650
            TabIndex        =   19
            Top             =   2670
            Width           =   2115
         End
         Begin VB.TextBox txtOChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   18
            Top             =   1320
            Width           =   2115
         End
         Begin VB.TextBox txtAbbrNm 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   1650
            TabIndex        =   17
            Top             =   3120
            Width           =   2115
         End
         Begin VB.TextBox txtResSpec 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   1650
            TabIndex        =   16
            Top             =   3570
            Width           =   1215
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   1650
            TabIndex        =   15
            Top             =   870
            Width           =   1245
         End
         Begin VB.TextBox txtRefLow 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   2280
            TabIndex        =   14
            Top             =   4020
            Width           =   1485
         End
         Begin VB.TextBox txtRefHigh 
            Alignment       =   2  '가운데 맞춤
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
            Height          =   300
            Left            =   2280
            TabIndex        =   13
            Top             =   4440
            Width           =   1485
         End
         Begin VB.CommandButton cmdSeqDown 
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3330
            TabIndex        =   12
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdSeqUp 
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2910
            TabIndex        =   11
            Top             =   840
            Width           =   405
         End
         Begin VB.CommandButton cmdSpecDown 
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3330
            TabIndex        =   10
            Top             =   3540
            Width           =   435
         End
         Begin VB.CommandButton cmdSpecUP 
            Caption         =   "▲"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2880
            TabIndex        =   9
            Top             =   3540
            Width           =   435
         End
         Begin VB.Frame frameCut 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1440
            TabIndex        =   28
            Top             =   5460
            Visible         =   0   'False
            Width           =   2565
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "사용"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   30
               Top             =   180
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.OptionButton optCutUse 
               BackColor       =   &H00FFFFFF&
               Caption         =   "미사용"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   210
               TabIndex        =   29
               Top             =   180
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1125
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "High"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   41
            Left            =   1680
            TabIndex        =   149
            Top             =   4530
            Width           =   375
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Low"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   40
            Left            =   1680
            TabIndex        =   148
            Top             =   4110
            Width           =   360
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   24
            Left            =   330
            Picture         =   "frmMain.frx":17357
            Top             =   4890
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "QC 채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   600
            TabIndex        =   125
            Top             =   4920
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "순번"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   600
            TabIndex        =   74
            Top             =   933
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과형식"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   600
            TabIndex        =   73
            Top             =   5310
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   14
            Left            =   330
            Picture         =   "frmMain.frx":17741
            Top             =   5280
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "CutOff"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   600
            TabIndex        =   72
            Top             =   5700
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   1
            Left            =   330
            Picture         =   "frmMain.frx":17B2B
            Top             =   5670
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   600
            TabIndex        =   71
            Top             =   1839
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   11
            Left            =   330
            Picture         =   "frmMain.frx":17F15
            Top             =   1809
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   0
            Left            =   330
            Picture         =   "frmMain.frx":182FF
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "장비코드"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   600
            TabIndex        =   70
            Top             =   480
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   2
            Left            =   330
            Picture         =   "frmMain.frx":186E9
            Top             =   1356
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "오더채널"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   600
            TabIndex        =   69
            Top             =   1386
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   3
            Left            =   330
            Picture         =   "frmMain.frx":18AD3
            Top             =   2262
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사코드"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   600
            TabIndex        =   68
            Top             =   2292
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   4
            Left            =   330
            Picture         =   "frmMain.frx":18EBD
            Top             =   2715
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사명"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   600
            TabIndex        =   67
            Top             =   2745
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   5
            Left            =   330
            Picture         =   "frmMain.frx":192A7
            Top             =   3168
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사약어"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   13
            Left            =   600
            TabIndex        =   66
            Top             =   3198
            Width           =   720
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   6
            Left            =   330
            Picture         =   "frmMain.frx":19691
            Top             =   3621
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "소수점"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   600
            TabIndex        =   65
            Top             =   3651
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   8
            Left            =   330
            Picture         =   "frmMain.frx":19A7B
            Top             =   4074
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "참고치"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   600
            TabIndex        =   64
            Top             =   4104
            Width           =   540
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   16
            Left            =   330
            Picture         =   "frmMain.frx":19E65
            Top             =   903
            Width           =   150
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   3
            Left            =   3990
            Top             =   8550
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "처방코드"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   63
            Top             =   8640
            Width           =   1125
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   2
            Left            =   3990
            Top             =   7140
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사저장"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   36
            Top             =   7230
            Width           =   1125
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   1
            Left            =   2580
            Top             =   7140
            Width           =   1335
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사삭제"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2700
            TabIndex        =   35
            Top             =   7230
            Width           =   1125
         End
         Begin VB.Label lblActionTest 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Refresh"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   2670
            TabIndex        =   33
            Top             =   8640
            Width           =   1125
         End
         Begin VB.Shape shpA 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   0
            Left            =   2580
            Top             =   8550
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "ex)10.00"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   3390
            TabIndex        =   32
            Top             =   5280
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "ex)10.00"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   3930
            TabIndex        =   22
            Top             =   3270
            Width           =   825
         End
      End
      Begin FPSpread.vaSpread spdTest 
         Height          =   9195
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Width           =   14325
         _Version        =   393216
         _ExtentX        =   25268
         _ExtentY        =   16219
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   28
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmMain.frx":1A24F
      End
   End
   Begin VB.Frame frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   930
      TabIndex        =   37
      Top             =   2370
      Visible         =   0   'False
      Width           =   20685
      Begin VB.CommandButton cmdIF 
         Caption         =   "IF 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11970
         TabIndex        =   116
         Top             =   8280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "병원정보설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   11910
         Style           =   1  '그래픽
         TabIndex        =   113
         Top             =   990
         Width           =   1965
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   10710
         TabIndex        =   59
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton optComType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4620
         TabIndex        =   58
         Top             =   450
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Frame frameTCP 
         BackColor       =   &H00FFFFFF&
         Caption         =   " TCP-IP 설정 "
         Height          =   7935
         Left            =   6480
         TabIndex        =   52
         Top             =   900
         Width           =   5325
         Begin VB.OptionButton optTCPType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Client"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   62
            Top             =   390
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optTCPType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3030
            TabIndex        =   61
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox txtTCPPort 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   57
            Top             =   1320
            Width           =   2445
         End
         Begin VB.TextBox txtTCPIP 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   56
            Top             =   930
            Width           =   2445
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   7
            Left            =   840
            Picture         =   "frmMain.frx":1B176
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   1110
            TabIndex        =   60
            Top             =   480
            Width           =   465
         End
         Begin VB.Shape shpTcp 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
         Begin VB.Label lblTcpSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "저장"
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
            Height          =   285
            Left            =   3300
            TabIndex        =   55
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Port"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   25
            Left            =   1110
            TabIndex        =   54
            Top             =   1395
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   15
            Left            =   840
            Picture         =   "frmMain.frx":1B560
            Top             =   1365
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "IP"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   24
            Left            =   1110
            TabIndex        =   53
            Top             =   990
            Width           =   180
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   10
            Left            =   840
            Picture         =   "frmMain.frx":1B94A
            Top             =   960
            Width           =   150
         End
      End
      Begin VB.Frame frameCom 
         BackColor       =   &H00FFFFFF&
         Caption         =   " RS-232 설정 "
         Height          =   7935
         Left            =   420
         TabIndex        =   38
         Top             =   870
         Width           =   5325
         Begin VB.ComboBox cboPort 
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
            ItemData        =   "frmMain.frx":1BD34
            Left            =   2190
            List            =   "frmMain.frx":1BD36
            TabIndex        =   51
            Top             =   390
            Width           =   2205
         End
         Begin VB.ComboBox cboBaudrate 
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
            ItemData        =   "frmMain.frx":1BD38
            Left            =   2190
            List            =   "frmMain.frx":1BD3A
            TabIndex        =   50
            Top             =   780
            Width           =   2205
         End
         Begin VB.ComboBox cboDatabit 
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
            ItemData        =   "frmMain.frx":1BD3C
            Left            =   2190
            List            =   "frmMain.frx":1BD3E
            TabIndex        =   49
            Top             =   1170
            Width           =   2205
         End
         Begin VB.ComboBox cboStartbit 
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
            Left            =   2190
            TabIndex        =   48
            Top             =   1590
            Width           =   2205
         End
         Begin VB.ComboBox cboStopbit 
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
            Left            =   2190
            TabIndex        =   47
            Top             =   2070
            Width           =   2205
         End
         Begin VB.ComboBox cboParity 
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
            ItemData        =   "frmMain.frx":1BD40
            Left            =   2190
            List            =   "frmMain.frx":1BD42
            TabIndex        =   46
            Top             =   2520
            Width           =   2205
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "DataBit"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   33
            Left            =   1110
            TabIndex        =   45
            Top             =   1290
            Width           =   645
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   23
            Left            =   840
            Picture         =   "frmMain.frx":1BD44
            Top             =   1260
            Width           =   150
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   22
            Left            =   840
            Picture         =   "frmMain.frx":1C12E
            Top             =   450
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "통신포트"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   32
            Left            =   1110
            TabIndex        =   44
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   21
            Left            =   840
            Picture         =   "frmMain.frx":1C518
            Top             =   855
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Baudrate"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   31
            Left            =   1110
            TabIndex        =   43
            Top             =   885
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   20
            Left            =   840
            Picture         =   "frmMain.frx":1C902
            Top             =   1695
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Start Bit"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   30
            Left            =   1110
            TabIndex        =   42
            Top             =   1725
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   19
            Left            =   840
            Picture         =   "frmMain.frx":1CCEC
            Top             =   2100
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Stop Bit"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   29
            Left            =   1110
            TabIndex        =   41
            Top             =   2130
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   225
            Index           =   18
            Left            =   840
            Picture         =   "frmMain.frx":1D0D6
            Top             =   2550
            Width           =   150
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Parity"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   1110
            TabIndex        =   40
            Top             =   2580
            Width           =   525
         End
         Begin VB.Label lblComSave 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "저장"
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
            Height          =   285
            Left            =   3300
            TabIndex        =   39
            Top             =   6960
            Width           =   1125
         End
         Begin VB.Shape shpCom 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            FillColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   3180
            Top             =   6810
            Width           =   1365
         End
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "파일"
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevelop 
         Caption         =   "개발처 : 010-3737-0551"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "설정"
      Begin VB.Menu mnuComm 
         Caption         =   "통신설정"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "검사설정"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "바코드사용"
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
      Begin VB.Menu mnuSave 
         Caption         =   "결과전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "적용결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   "기타"
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objMSHTML As New MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument

Private Sub chkAll_Click()
    Dim iRow As Long
    
    With spdOrder
        If chkAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 1
            Next iRow
        ElseIf chkAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 0
            Next iRow
        End If
    End With
End Sub

Private Sub chkRAll_Click()
    Dim iRow As Long
    
    With spdROrder
        If chkRAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 1
            Next iRow
        ElseIf chkRAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 0
            Next iRow
        End If
    End With
End Sub

Private Sub cmdAnalyteFind_Click()

    frmQCList.Tag = "Analyte조회"
    DoEvents
    frmQCList.Show 'vbModal
    
End Sub

'Private Sub cmdRefresh_Click()
'
'    Call GetTestList
'
'End Sub

Private Sub cmdAppend_Click()

    spdOrdMst.MaxRows = spdOrdMst.MaxRows + 1
    
End Sub

Private Sub cmdConfig_Click()
    
    frmHospInfo.Show 'vbModal
    
End Sub

Private Sub cmdDelete_Click()
    
    spdOrdMst.Row = spdOrdMst.ActiveRow
    spdOrdMst.Action = ActionDeleteRow
    
    spdOrdMst.MaxRows = spdOrdMst.MaxRows - 1
    
End Sub

Private Sub cmdEnd_Click()

    If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
    
        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    End If
    
End Sub

Private Sub cmdIF_Click()

    If FraHidden.Visible = True Then
        FraHidden.Visible = False
    Else
        FraHidden.Visible = True
        FraHidden.ZOrder 0
    End If
    
End Sub

Private Sub cmdInit_Click()
    
    Call InitialComm
    
End Sub

Private Sub cmdLabFind_Click()

    frmQCList.Caption = "Lab조회"
    frmQCList.Show 'vbModal
    
End Sub

Private Sub cmdQCMaster_Click()

    frmQCMaster.Show 'vbModal
    
End Sub

Private Sub cmdResult_Click()

    Call cmdGo_Click
    
    Call cmdParse_Click
    
    Call cmdReadJSON_Click

End Sub

Private Sub cmdGo_Click()
'    Screen.MousePointer = vbHourglass
    
    txtLink = ""
    txtSource = ""
    'cmdGo.Enabled = False
    'cmdParse.Enabled = False
    
    lblStatus.Caption = "Gettting document via HTTP"
    
    ' This function is only available with Internet Explorer 5
    
    Set objDocument = objMSHTML.createDocumentFromUrl(txtURL.Text, vbNullString)
    
    Call SetSQLData("objDocument", objDocument.url & vbCr & objDocument.Title & vbCr & objDocument.fileModifiedDate & vbCr & objDocument.fileSize)
    
    lblStatus.Caption = "Getting and parsing HTML document"
    
    ' Tricky, to make the function wait for the document to complete, usually
    ' the transfer is asynchronus. Note that this string might be different if
    ' you have another language than english for Internet Explorer on the
    ' machine where the code is executed.
    Dim tm As Date, ovtm As Boolean
    
    tm = Now
    'cmdGo.Enabled = False
    Do While objDocument.readyState <> "complete"
        If Now - tm > 10# / 86400 Then
            ovtm = True
            Exit Do
        End If
        DoEvents
    Loop
    
    lblStatus.Caption = "Document completed"
    
    txtSource.Text = objDocument.documentElement.outerHTML
    
    ' Copying the title of the page to the label
    
    lblTitle.Caption = "Title : " & objDocument.Title
    txtLink.Text = ""
    
    txtLink.SelText = "URL = " & objDocument.url & vbCrLf
    txtLink.SelText = "Title = " & objDocument.Title & vbCrLf
    txtLink.SelText = "Modified = " & objDocument.fileModifiedDate & vbCrLf
    txtLink.SelText = "Size = " & objDocument.fileSize & vbCrLf
    
    If ovtm Then
        lblStatus.Caption = "Done, but overtime, so may be not fully parsed."
    Else
        lblStatus.Caption = "Done"
    End If
    
    'cmdGo.Enabled = True
    'cmdParse.Enabled = True
    'cmdExit.Enabled = True
    
    
'    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdParse_Click()
    Dim objLink As HTMLLinkElement
    Dim obj As IHTMLElement
    Dim emb As HTMLEmbed
    Dim frm As HTMLFormElement
    Dim img As HTMLImg
    'Dim Span As HTMLSpanElement
    Dim s As String
    Dim v As Variant
   ' Dim colObj As New Collection
   ' Dim obj1 As Object
    
    'On Error Resume Next
    
    ' Copying the source to the text box
'    Screen.MousePointer = vbHourglass
    'cmdGo.Enabled = False
    'cmdParse.Enabled = False
'    cmdExit.Enabled = False
    
    lblStatus.Caption = "Extracting links"
    
    ' Processing the link collection of the HTMLDocument object
    Dim l As String, p As Long
    
    l = Trim(objDocument.location)
    
    'If InStr(l, "?") Then
        p = InStrRev(l, "/")
        If p Then
            l = Left$(l, p)
        End If
    'End If
    
    
    txtLink.Text = ""
    txtLink.SelText = "URL = " & objDocument.url & vbCrLf
    txtLink.SelText = "Title = " & objDocument.Title & vbCrLf
    txtLink.SelText = "Modified = " & objDocument.fileModifiedDate & vbCrLf
    txtLink.SelText = "Size = " & objDocument.fileSize & vbCrLf
    
    If chkFullLink = 1 Or chkLink = 1 Then
        txtLink.SelText = vbCrLf & "<links>" & vbCrLf
        For Each objLink In objDocument.links
            s = Trim(objLink)
            If (chkFullLink = 1) Or (Left(s, Len(l)) <> l) Then
                txtLink.SelText = s & vbCrLf
                lblStatus.Caption = "Extracted " & objLink
                lblStatus.Refresh
                DoEvents
            End If
        Next
        
        txtLink.SelText = vbCrLf & "<OnClick>" & vbCrLf
        For Each obj In objDocument.All 'anchors '.documentElement
            If obj Is Nothing Then
            ElseIf obj.children.Length > 1 Then
            ElseIf IsNull(obj.onclick) Then
            ElseIf Len(Trim(obj.onclick)) Then
                v = Split(Trim(obj.onclick), """")
                If UBound(v) > 0 Then
                    s = Trim(v(1))
                    p = InStr(UCase(s), "SRC=")
                    If p Then
                        s = Mid$(s, p + 4)
                    End If
                    If chkFullLink Then
                        txtLink.SelText = "<" & obj.tagName & "> " & s & vbCrLf
                    Else
                        txtLink.SelText = s & vbCrLf
                    End If
                ElseIf chkFullLink Then
                    v = Split(Trim(obj.onclick), "'")
                    If UBound(v) > 0 Then
                        txtLink.SelText = "<" & obj.tagName & "> " & v(1) & vbCrLf
                    Else
                        v = Split(Trim(obj.onclick), Chr(10))
                        If UBound(v) > 0 Then
                            txtLink.SelText = "<" & obj.tagName & "> " & v(UBound(v) - 1) & vbCrLf
                        Else
                            txtLink.SelText = "<" & obj.tagName & "> " & v(0) & vbCrLf
                        End If
                    End If
                End If
                DoEvents
            End If
        Next
    End If
    
    If chkImbed = 1 Then
        txtLink.SelText = vbCrLf & "<embeds>" & vbCrLf
        For Each emb In objDocument.embeds
            txtLink.SelText = emb.src & vbCrLf
            lblStatus.Caption = "Extracted " & emb.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkForms = 1 Then
        txtLink.SelText = vbCrLf & "<Forms>" & vbCrLf
        For Each frm In objDocument.Forms
            txtLink.SelText = frm.Name & vbCrLf
            lblStatus.Caption = "Extracted " & frm.Name
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkImage = 1 Then
        txtLink.SelText = vbCrLf & "<images>" & vbCrLf
        For Each img In objDocument.images
            txtLink.SelText = img.src & vbTab & vbTab & img.fileModifiedDate & vbCrLf
            lblStatus.Caption = "Extracted " & img.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkPlugin = 1 Then
        txtLink.SelText = vbCrLf & "<plugins>" & vbCrLf
        For Each obj In objDocument.plugins
            txtLink.SelText = obj.src & vbCrLf
            lblStatus.Caption = "Extracted " & obj.src
            lblStatus.Refresh
            DoEvents
        Next
    End If
    
    If chkScript = 1 Then
        txtLink.SelText = vbCrLf & "<scripts>" & vbCrLf
        For Each obj In objDocument.scripts
            s = Trim(obj.src)
            If Len(s) Then
                txtLink.SelText = s & vbCrLf
                lblStatus.Caption = "Extracted " & s
            ElseIf Len(Trim(obj.Title)) Then
                txtLink.SelText = obj.Title & vbCrLf
            ElseIf Len(Trim(obj.className)) Then
                txtLink.SelText = obj.className & vbCrLf
            ElseIf Len(Trim(obj.innerHTML)) Then
                s = Trim(obj.innerHTML)
                v = Split(s, vbLf)
                txtLink.SelText = "Gloval" & vbTab '& vbCrLf
                Dim v1 As Variant
                Dim isFun As Boolean, isCmt As Boolean
                For Each v1 In v
                    If Not isCmt And Left$(v1, 2) = "//" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isCmt = True
                    ElseIf Not isFun And Left$(v1, 3) = "var" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                    ElseIf Left$(v1, 3) = "fun" Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isFun = True
                    ElseIf InStr(v1, vbTab & "function ") Then
                        s = vbTab & v1 & vbCrLf
                        txtLink.SelText = s
                        isFun = True
                    End If
                Next
                
            ElseIf Len(Trim(obj.innerText)) Then
                txtLink.SelText = obj.innerText & vbCrLf
            ElseIf Len(Trim(obj.tagurl)) Then
                txtLink.SelText = obj.tagurl & vbCrLf
            Else
                txtLink.SelText = "Blank" & vbCrLf
            End If
            lblStatus.Refresh
            DoEvents
        Next
    End If
            
'    cmdGo.Enabled = True
'    cmdParse.Enabled = True
'    cmdExit.Enabled = True
'    Screen.MousePointer = vbNormal
End Sub

'''{"UserMetadata":
'''                                {"AssayRunInFactoryMode":false,
'''                                "Category":0,
'''                                "PatientId":"qw",
'''                                "SampleType":"Vtm",
'''                                "Cartridge":{
'''                                                        "BarcodeId":"02000000M079262",
'''                                                        "ProductId":"02",
'''                                                        "LotNumber":"M079262",
'''                                                        "CartridgeBarcodeResult":1,
'''                                                        "SerialNumber":"",
'''                                                        "IsValid":true},
'''                                "UserId":"admin",
'''                                "Comments":null},
'''
'''"RawData":null,
'''"Decision":
'''                                {"TestResults":
'''                                                            {"$type":"System.Collections.Generic.Dictionary`2[[System.String, mscorlib],[Iso.Domain.AssayDecision.Enums.TestResultEnum, Iso.Domain]], mscorlib",
'''                                                            "FluA":0,
'''                                                            "FluB":0},
'''                                "ChannelResults":
'''                                                            {"$type":"System.Collections.Generic.Dictionary`2[[System.String, mscorlib],[Iso.Domain.AssayDecision.Enums.ChannelResult, Iso.Domain]], mscorlib"},
'''                                "EvaluationCriteriaResults":
'''                                                            {"$type":"System.Collections.Generic.Dictionary`2[[System.String, mscorlib],[System.Collections.Generic.IDictionary`2[[System.Int32, mscorlib],[Iso.Domain.AssayDecision.CriteriaEvaluationResult, Iso.Domain]], mscorlib]], mscorlib"},
'''                                "CategoryType":0,"Passed":false,"ProceduralControlValid":false},
'''                                                            "Definition":{"Name":"Flu A\/B",
'''                                                            "Description":"Influenza A & B Test",
'''                                                            "TestCodeId":"02",
'''                                                            "Heaters":
'''                                                                                {"HeaterTemperatures":
'''                                                                                                    {"$type":"System.Collections.Generic.Dictionary`2[[Iso.Domain.Enums.HeaterType, Iso.Domain],[System.Single, mscorlib]], mscorlib",
'''                                                                                                    "PreHeater":56,"RotorHeater":56}},
'''                                                                                                    "CommandList":null,
'''                                                                                                    "Tubes":3,
'''                                                                                                    "OptoTypes":["FAM","ROX"],
'''                                                                                                    "UniqueId":"00000000-0000-0000-0000-000000000000",
'''                                                                                                    "Id":null,
'''                                                                                                    "StartedTimestamp":"\/Date(-62135596800000+0000)\/",
'''                                                                                                    "ValidationValue":null,
'''                                                                                                    "LinkedTestType":null},
'''                                                                                                    "UniqueId":"c67d764c-46fd-4c05-82a9-ca8276f94709",
'''                                                                                                    "StartedTimestamp":"\/Date(1456227814248+0000)\/",
'''                                                                                                    "EndedTimestamp":"\/Date(1456228416014+0000)\/",
'''                                                                                                    "InstrumentDetails":
'''                                                                                                                                        {"InstrumentSerialNumber":"A1F93616",
'''                                                                                                                                        "HardwareVersion":"4",
'''                                                                                                                                        "ApplicationVersion":"4.0.5.4.187-pro",
'''                                                                                                                                        "DeviceMacAddress":"00:04:A3:E5:F5:2C",
'''                                                                                                                                        "KernelVersion":"3.0.4r22-hg684c0d9ff0a7-b129+ ",
'''                                                                                                                                        "UpdateNumber":"0411",
'''                                                                                                                                        "MercurialRevision":"43483e2037b4",
'''                                                                                                                                        "RfsVersion":" b129-ts2014-09-24-17-08-30-hg7a2bb8a8afef",
'''                                                                                                                                        "FamOptoVersion":"3.0.8",
'''                                                                                                                                        "RoxOptoVersion":"3.0.8",
'''                                                                                                                                        "TrayVersion":"3.8.3",
'''                                                                                                                                        "BaseboardVersion":"2.00",
'''                                                                                                                                        "BaseboardHWVersion":"3",
'''                                                                                                                                        "ConfigurationName":"B.1.0.0"},
'''                                                                                                                                        "ValidationValue":"3D1F0110615AD48391D984E804ABFC05",
'''                                                                                                                                        "RunState":2,
'''                                                                                                                                        "Id":null}


Private Sub cmdReadJSON_Click()

    Dim p As Object
   
    Dim varLink     As Variant
    Dim strFilePath As String
    Dim strJSON     As String
    Dim i As Long
    Dim j As Long
    Dim strHTML As String
    
    'On Error Resume Next
    
    varLink = Split(txtLink.Text, vbCrLf)
    Erase strRecvData
    j = 0
    txtHTML.Text = ""
    
    'MsgBox UBound(varLink)
    
    'Screen.MousePointer = 11
    
    For i = 0 To UBound(varLink)
        If Mid(varLink(i), 1, 7) = "http://" And Right(varLink(i), 4) = "json" Then
            Debug.Print varLink(i)
            ReDim Preserve strRecvData(j)

            strFilePath = CStr(varLink(i))
            
            'MsgBox strFilePath
            
            '-- 테스트용 : 컴파일시 제외할것 !!!!!!!!!!        Alerei_Patient_02_34507017_16-12-03_09-25-01_b222b707-afbe-4151-af67-31fea613c3ba
            'strFilePath = "D:\프로젝트\VB\강안병원_자인컴\Test\Alerei_Patient_02_34507017_16-12-03_09-25-01_b222b707-afbe-4151-af67-31fea613c3ba.json"
            '-- 테스트용 : 컴파일시 제외할것 !!!!!!!!!!
            
            strJSON = ""
            strJSON = ParserJSON(strFilePath)
            
            '-- 2016.12.08 추가 ParserJSON에서 데이터를 못 찾아오는 오류로 인해 추가함
            strHTML = Inet1.OpenURL(strFilePath)
            txtHTML.Text = strHTML
            Text1.Text = Text1 & txtHTML.Text
            strJSON = txtHTML.Text
            '-------------------------------------------------------------------------
            
            If InStr(strJSON, "{") > 0 Then
                strJSON = Mid(strJSON, InStr(strJSON, "{"))
                strJSON = Mid(strJSON, 1, Len(strJSON) - 20)
            End If
            
            If Mid(strJSON, 1, 1) = "{" And InStr(strFilePath, "index.") <= 0 And InStr(strFilePath, "_QC_") <= 0 Then
                Set p = JSON.parse(strJSON)
    
                If Not (p Is Nothing) Then
                    'MsgBox JSON.GetParserErrors
                    'If JSON.GetParserErrors <> "" Then
                        'MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
                    'Else
                        'MsgBox "Base item count: " & p.Count
                        'MsgBox "JSON toString: " & Left(JSON.toString(p), 1000)
                        
    '                    Debug.Print p.Item("UserMetadata").Item("PatientId")
    '                    Debug.Print p.Item("UserMetadata").Item("SampleType")
    '                    Debug.Print p.Item("UserMetadata").Item("Cartridge").Item("BarcodeId")
    '                    Debug.Print p.Item("UserMetadata").Item("Cartridge").Item("LotNumber")
    '                    Debug.Print p.Item("UserMetadata").Item("UserId")
    '                    Debug.Print p.Item("Decision").Item("TestResults").Item("FluA")
    '                    Debug.Print p.Item("Decision").Item("TestResults").Item("FluB")
                        
                                         strRecvData(j) = CStr(p.Item("UserMetadata").Item("PatientId")) & "|"
                        strRecvData(j) = strRecvData(j) & CStr(p.Item("Decision").Item("TestResults").Item("FluA")) & "|"
                        strRecvData(j) = strRecvData(j) & CStr(p.Item("Decision").Item("TestResults").Item("FluB")) & "|"
                        
                        
                    'End If
                Else
                    'MsgBox "An error occurred parsing " & strFilePath
                End If
                
                'SetRawData "[Rx]" & "Base item count: " & p.Count
                SetRawData "[Rx]" & j & ">>" & strFilePath & ":" & strRecvData(j)
                j = j + 1
            End If
        End If
    Next
    
    'Screen.MousePointer = 0
   
    If UBound(strRecvData) > 0 Then
        Call HTMLRcvData_ALLEREi
    End If
    
    Exit Sub
    
'    If cd.Filename <> "" Then
'       Set p = JSON.parse(ReadTextFile(cd.Filename))
'       If Not (p Is Nothing) Then
'          If JSON.GetParserErrors <> "" Then
'             MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
'          Else
'             MsgBox "Base item count: " & p.Count
'             MsgBox "JSON toString: " & Left(JSON.toString(p), 1000)
'          End If
'       Else
'          MsgBox "An error occurred parsing " & cd.Filename
'       End If
'    End If
   
End Sub

Private Function ParserJSON(ByVal JsonURL As String) As String
'    Screen.MousePointer = vbHourglass
    Dim tm As Date, ovtm As Boolean
    
    Set objDocument = objMSHTML.createDocumentFromUrl(JsonURL, vbNullString)
  
    tm = Now
    cmdGo.Enabled = False
    Do While objDocument.readyState <> "complete"
        If Now - tm > 10# / 86400 Then
            ovtm = True
            Exit Do
        End If
        DoEvents
    Loop
    
    ParserJSON = objDocument.documentElement.outerHTML
    
    Set objDocument = Nothing
    
End Function

Private Sub cmdSend_Click()
    Dim i As Integer
    Dim varTmp As Variant
    
    Erase strRecvData
    varTmp = Replace(txtRcv.Text, vbLf, "")
    varTmp = Split(varTmp, vbCr)
    
    For i = 0 To UBound(varTmp)
        ReDim Preserve strRecvData(i + 1)
        strRecvData(i + 1) = varTmp(i)
    Next
    
    Select Case UCase(gHOSP.MACHNM)
        ' 타이머를 사용해야 해서 메인에서 처리함
        Case "ADVIA2120-1", "ADVIA2120-2"
                Call Phase_Serial_ADVIA2120
                
        ' 타이머를 사용해야 해서 메인에서 처리함
        Case "CT500"
                Call Phase_Serial_CT500
                
        Case "VERSACELL"
                Call Phase_Serial_VERSACELL
        
        Case "RAPIDLAB348"
                Call Phase_Serial_RAPIDLAB348
        
        Case "PFA200"
                Call Phase_Serial_PFA200
        
        Case "AFIAS6"
                Call Phase_Serial_AFIAS6
        
        Case "ADVIA1800-1", "ADVIA1800-2"
                Call Phase_Serial_ADVIA1800
        
        Case "RAPIDPOINT500"
                Call Phase_Serial_RAPIDPOINT500
        
        Case "ACLTOP"
                Call Phase_Serial_ACLTOP
        
        Case "VESCUBE"
                Call Phase_Serial_VESCUBE
        
        Case "OSMOPRO"
                Call Phase_Serial_OSMOPRO
        
        Case "ALEREI"
             Call TCPRcvData_ALEREi
        Case Else
                Call Serial_Protocol
                
    End Select




End Sub

Private Sub cmdSeqDown_Click()
    On Error Resume Next
    
    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    On Error Resume Next
    
    txtSeq.Text = txtSeq.Text + 1

End Sub

Private Sub cmdSet_Click()

    If frameSet.Visible = True Then
        frameSet.Visible = False
        
    Else
        frameSet.Visible = True
        frameSet.ZOrder 0
    End If
    
End Sub

Private Sub cmdSL_Click()

    If cmdSL.Caption = "▶" Then
        cmdSL.Caption = "◀"
        spdOrder.Width = Me.Width - 400
    Else
        cmdSL.Caption = "▶"
        spdOrder.Width = Me.ScaleWidth - spdResult.Width - 280
    End If
    
End Sub

Private Sub cmdSpecDown_Click()
    On Error Resume Next
    
    txtResSpec.Text = txtResSpec.Text - 1

End Sub

Private Sub cmdSpecUP_Click()
    On Error Resume Next
    
    txtResSpec.Text = txtResSpec.Text + 1

End Sub


Private Sub cmdTVSave_Click()
    Dim i As Integer
    
    If Trim(txtTV.Text) <> "" Then
        If IsNumeric(Trim(txtTV.Text)) Then
            If lblPatInfo(1).Caption <> "" And lblPatInfo(2).Caption <> "" Then
                Call SetLocalDB_TV(spdROrder.ActiveRow, 1, 1, txtTV.Text)
                
                Call spdROrder_Click(colBARCODE, spdROrder.ActiveRow)
                        
                For i = 1 To spdRResult.MaxRows
                    Call CalProcess(spdROrder, spdRResult, GetText(spdRResult, i, colRTESTCD), Trim(txtTV.Text))
                Next
            End If
        Else
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbCritical, Me.Caption
            txtTV.SetFocus
        End If
    End If

End Sub

Private Sub cmdView_Click()

    If fraSetup.Visible = True Then
        fraSetup.Visible = False
        fraLinks.Visible = False
        fraSource.Visible = False
    Else
        fraSetup.Visible = True
        fraLinks.Visible = True
        fraSource.Visible = True
        
        fraSetup.ZOrder 0
        fraLinks.ZOrder 0
        fraSource.ZOrder 0
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call cmdEnd_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'If MsgBox("장비와 통신중입니다. 종료하시겠습니까?", vbYesNo + vbCritical, "프로그램 종료") = vbYes Then
    
        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    'End If
    
End Sub

Private Sub fraResult_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRS.BorderColor = &H808080

End Sub



Private Sub lblRClear_Click()
    
    spdROrder.MaxRows = 0
    spdRResult.MaxRows = 0
    
End Sub

Private Sub lblRClear_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
    lblRClear.ForeColor = vbBlue
    shpRC.BorderColor = vbCyan
    
End Sub

Private Sub lblResult_Click()

    frmMain.spdROrder.MaxRows = 0
    frmMain.spdRResult.MaxRows = 0

    Call GetResultList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), cboRstType.ListIndex, cboState.ListIndex)
    
End Sub

Private Sub lblResult_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRS.BorderColor = &H808080

    
    lblResult.ForeColor = vbBlue
    shpR.BorderColor = vbCyan
    
End Sub

Private Sub lblRSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    For lRow = 1 To spdROrder.DataRowCnt
        spdROrder.Row = lRow
        spdROrder.Col = 1
        If spdROrder.Value = 1 Then
            
            Res = SaveTransDataR_BDJS(lRow)
        
            If Res = -1 Then
                SetForeColor spdROrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                SetText spdROrder, "Failed", lRow, colSTATE
            Else
                spdROrder.Row = lRow
                spdROrder.Col = 1
                spdROrder.Value = 1
                
                SetBackColor spdROrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                SetText spdROrder, "Trans", lRow, colSTATE
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdROrder, lRow, colBARCODE)) & "' "
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
                
            End If
            spdROrder.Row = lRow
            spdROrder.Col = 1
            spdROrder.Value = 0
        End If
    Next lRow
    

'    Dim lRow As Long
'    Dim Res  As Integer
'
'    For lRow = 1 To spdOrder.DataRowCnt
'        spdOrder.Row = lRow
'        spdOrder.Col = 1
'        If spdOrder.Value = 1 Then
'
'            Res = SaveTransData_BDJS(lRow)
'
'            If Res = -1 Then
'                SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
'                SetText spdOrder, "Failed", lRow, colSTATE
'            Else
'                spdOrder.Row = lRow
'                spdOrder.Col = 1
'                spdOrder.Value = 1
'
'                SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
'                SetText spdOrder, "Trans", lRow, colSTATE
'
'                      SQL = " UPDATE PATRESULT SET " & vbCrLf
'                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
'                SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
'                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
'
'                If DBExec(AdoCn_Local, SQL) Then
'                    '-- 성공
'                End If
'
'            End If
'            spdOrder.Row = lRow
'            spdOrder.Col = 1
'            spdOrder.Value = 0
'        End If
'    Next lRow
End Sub

Private Sub lblRSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblResult.ForeColor = vbBlack
    lblRClear.ForeColor = vbBlack
    lblRSave.ForeColor = vbBlack
    
    shpR.BorderColor = &H808080
    shpRC.BorderColor = &H808080
    shpRS.BorderColor = &H808080
    
    lblRSave.ForeColor = vbBlue
    shpRS.BorderColor = vbCyan

End Sub

Private Sub lblSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    For lRow = 1 To spdOrder.DataRowCnt
        spdOrder.Row = lRow
        spdOrder.Col = 1
        If spdOrder.Value = 1 Then
            
            Res = SaveTransData_BDJS(lRow)
        
            If Res = -1 Then
                SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                SetText spdOrder, "Failed", lRow, colSTATE
            Else
                spdOrder.Row = lRow
                spdOrder.Col = 1
                spdOrder.Value = 1
                
                SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                SetText spdOrder, "Trans", lRow, colSTATE
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
                
            End If
            spdOrder.Row = lRow
            spdOrder.Col = 1
            spdOrder.Value = 0
        End If
    Next lRow

End Sub

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuComm_Click()
    
    Call lblMenu_Click(3)

End Sub

Private Sub mnuCommTest_Click()

    If fraCommTest.Visible = False Then
        fraCommTest.Visible = True
    Else
        fraCommTest.Visible = False
    End If
    
End Sub

Private Sub mnuEqpResult_Click()
    
    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub


Private Sub mnuExit_Click()
    Call cmdEnd_Click
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHelp02_Click()

    Call WinExec("C:\Program Files (x86)\Internet Explorer\iexplore.exe http://cs1472.com/customer/", 1)

End Sub

Private Sub mnuLisResult_Click()
    
    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuRackPos_Click()
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuSaveAuto_Click()
    
    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuSaveManual_Click()
    
    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True
    
    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")


End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False
    
    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")

End Sub

Private Sub mnuTest_Click()
    
    Call lblMenu_Click(2)

End Sub

'Private Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim sRow        As Long
'    Dim sCol        As Long
'    Dim strNewBarNo As String
'
'    sRow = spdOrder.ActiveRow
'    sCol = spdOrder.ActiveCol
'    strNewBarNo = GetText(spdOrder, sRow, sCol)
'    'MsgBox sRow
'    'MsgBox sCol
'
'    If KeyCode = vbKeyReturn Then
'        If colBARCODE = sCol Then
'            If GetSampleInfo(sRow, spdOrder) = -1 Then
'                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
'            Else
'                '정보수정
'                SQL = ""
'                SQL = SQL & "UPDATE PATRESULT SET "
'                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCr
'                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCr
'                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCr
'                SQL = SQL & " ,PID     = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCr
'                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCr
'                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCr
'                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCr
'                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdOrder, sRow, colPJUMIN)) & "'" & vbCr
'                SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCr
'                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCr
'                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
'
'                If DBExec(AdoCn_Local, SQL) Then
'                    '-- 성공
'                End If
'            End If
'        End If
'    End If
'
'End Sub


Private Sub spdOrder_KeyPress(KeyAscii As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    strNewBarNo = GetText(spdOrder, sRow, sCol)
    'MsgBox sRow
    'MsgBox sCol
    
    If KeyAscii = vbKeyReturn Then
        If colBARCODE = sCol Then
            If GetSampleInfo(sRow, spdOrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCr
                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCr
                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCr
                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdOrder, sRow, colPJUMIN)) & "'" & vbCr
                SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
        End If
    End If
End Sub

Private Sub spdROrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    If Row = 0 Then
        Call SetSpreadSort(spdROrder, 0)
        Exit Sub
    End If
    
    '-- 환자정보표시
    lblPatInfo(0).Caption = GetText(spdROrder, Row, colPNAME) & " [" & GetText(spdROrder, Row, colPAGE) & "/" & GetText(spdROrder, Row, colPSEX) & "]  "
    lblPatInfo(1).Caption = GetText(spdROrder, Row, colBARCODE)
    lblPatInfo(2).Caption = GetText(spdROrder, Row, colPID)
    lblPatInfo(3).Caption = spdROrder.ActiveRow
    lblPatInfo(4).Caption = GetText(spdROrder, Row, colRACKNO)

    
    txtTV.Text = ""
    
    '-- 결과표시
    If GetPatTRestResult_Search(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdRResult.MaxRows = 0
        With spdROrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdROrder, Row, intCol) <> "" Then    '◆
                    spdRResult.MaxRows = spdRResult.MaxRows + 1
                    Call SetText(spdRResult, GetText(spdROrder, 0, intCol), spdRResult.MaxRows, colRTESTNM)
                    spdRResult.RowHeight(-1) = 12
                End If
            Next
        End With
    End If

    'txtTV.SetFocus
    
End Sub

Private Sub spdROrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim intRow      As Long
    Dim strTestCd   As String
    Dim strTestNm   As String
    Dim strResult   As String
    Dim strIntBase  As String
    Dim strJudge    As String
    Dim lsID        As String
    Dim lsSeq       As Long
    Dim strExamDate As String
    
    sRow = spdROrder.ActiveRow
    sCol = spdROrder.ActiveCol
    
    If KeyCode = vbKeyDelete Then
        If sRow < 1 Or sRow > spdROrder.DataRowCnt Then
            Exit Sub
        End If
        If sCol > colSTATE Then
            Exit Sub
        End If
        lsID = Trim(GetText(spdROrder, sRow, colBARCODE))
        lsSeq = Trim(GetText(spdROrder, sRow, colSAVESEQ))
        strExamDate = Trim(GetText(spdROrder, sRow, colEXAMDATE))
        

        If lsSeq < 1 Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " 의 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        'SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & strExamDate & "' "
        
'        Res = SendQuery(gLocal, SQL)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- 성공
        End If
                
        DeleteRow spdROrder, sRow, sRow
        spdRResult.MaxRows = 0
        'blnModify = True
        
    ElseIf KeyCode = vbKeyReturn Then
        If spdROrder.ActiveCol = colBARCODE Then
            
            If GetSampleInfo(sRow, spdROrder) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '-- 환자정보수정
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE = '" & Trim(GetText(spdROrder, sRow, colBARCODE)) & "'" & vbCr
                SQL = SQL & " ,INOUT   = '" & Trim(GetText(spdROrder, sRow, colINOUT)) & "'" & vbCr
                SQL = SQL & " ,CHARTNO = '" & Trim(GetText(spdROrder, sRow, colCHARTNO)) & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(spdROrder, sRow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(spdROrder, sRow, colPNAME)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(spdROrder, sRow, colPSEX)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(spdROrder, sRow, colPAGE)) & "'" & vbCr
                SQL = SQL & " ,PJUMIN  = '" & Trim(GetText(spdROrder, sRow, colPJUMIN)) & "'" & vbCr
                SQL = SQL & " WHERE mid(EXAMDATE,1,8) = '" & Mid(Trim(GetText(spdROrder, sRow, colEXAMDATE)), 1, 8) & "'" & vbCr
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdROrder, sRow, colSAVESEQ)) & vbCr
                'SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
                'SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdROrder, asRow1, colBARCODE)) & "' " & vbCr
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            
        ElseIf spdROrder.ActiveCol > colSTATE Then
            strTestNm = GetText(spdROrder, 0, sCol)
            strResult = GetText(spdROrder, sRow, sCol)
            
            For intRow = 1 To spdRResult.MaxRows
                If strTestNm = GetText(spdRResult, intRow, colRTESTNM) Then
                    strTestCd = GetText(spdRResult, intRow, colRTESTCD)
                    strIntBase = GetText(spdRResult, intRow, colRCHANNEL)
                
                    '소수점 처리, 결과판정
                    strResult = SetResult(strResult, strIntBase)
                    strJudge = SetJudge(strResult, strIntBase)
                                                        
                                                        
                    '-- 검사결과수정
                    SQL = ""
                    SQL = SQL & "UPDATE PATRESULT SET "
                    SQL = SQL & "  RESULT   = '" & strResult & "'" & vbCr
                    SQL = SQL & " ,REFJUDGE = '" & strJudge & "'" & vbCr
                    SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(spdROrder, sRow, colEXAMDATE)) & "'" & vbCr
                    SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdROrder, sRow, colSAVESEQ)) & vbCr
                    SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = '" & strTestCd & "'" & vbCr
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                        Call SetText(spdROrder, strResult, sRow, sCol)
                        Call spdROrder_Click(sCol, sRow)
                    End If
                End If
            Next
        End If
    End If
End Sub

Private Sub InitialComm()
    Dim sSendBuf$
    
    intPhase = 1
    
    msMT = Chr(&H30)
    sSendBuf = msMT & "I " & vbCr & vbLf
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    comEqp.Output = sSendBuf
    SetRawData "[Tx]" & sSendBuf
       
End Sub

Private Sub Timer2120_Timer()
    Timer2120.Enabled = False
    Timer2120.Interval = 0
    
    Select Case msTimerFlag
        Case "I"
            If msMT = "" Then
                Call InitialComm
                
                msSndPacket = ""
                msTimerFlag = ""
            Else
                mp_bReserveEnd = True
                'PropertyChanged "ReserveEnd"
                
                If mp_bPortOpen = False Then
                    mp_bReserveEnd = False
                    'PropertyChanged "ReserveEnd"
                    
                    comEqp.PortOpen = True
                    mp_bPortOpen = True
                    'PropertyChanged "PortOpen"
                    
                    Sleep 1000
                    
                    Call InitialComm
                    
                    msSndPacket = ""
                    msTimerFlag = ""
                Else
                    'Timer 재가동
                    msTimerFlag = "I"
                    Timer2120.Interval = 1000
                    Timer2120.Enabled = True
                End If
            End If
            
        Case Else
            If msSndPacket = "" Then Exit Sub
    
            comEqp.Output = msSndPacket
            SetRawData "[Tx]" & msSndPacket

            msSndPacket = ""
            msTimerFlag = ""
            
    End Select
End Sub

Private Sub TimerVESCUBE_Timer()
    
    TimerVESCUBE.Enabled = False
    
    Call VesMatic(RcvBuffer)
    
End Sub

Private Sub tmrComm_Timer()
    
    tmrComm.Enabled = False
    tmrFlipFlop.Enabled = False

    
    lblCommStatus.Caption = ""
    
End Sub

Private Sub tmrFlipFlop_Timer()
    lblCommStatus.ForeColor = vbBlue
    
    If lblCommStatus.Visible = True Then
        lblCommStatus.Visible = False
    Else
        lblCommStatus.Visible = True
    End If
    
    
    If lblMenu(0).ForeColor = vbBlack Then
        lblMenu(0).ForeColor = vbBlue
        shpB(0).BorderColor = vbCyan
    Else
        lblMenu(0).ForeColor = vbBlack
        shpB(0).BorderColor = vbGreen
    End If

End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrResult_Timer()

    If gComm.TCPTYPE = "1" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    tmrResult.Enabled = False

'
'    txtTimer.Text = txtTimer.Text - 1
'
'    If txtTimer.Text = "0" Then
'        Call cmdResult_Click
'        txtTimer.Enabled = False
'        txtTimer.Text = "60"
'    End If
'
End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

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
Private Sub GetOrder_ADVIA2120(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
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
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ADVIA2120(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub



'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ADVIA2120(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String
'    Dim blnCBC      As Boolean
'    Dim blnDIFF     As Boolean
'    Dim blnRETI     As Boolean
'    Dim strPART     As String

    GetEquipExamCode_ADVIA2120 = ""
    strExamCode = ""
    mOrder.SendCnt = 0
    
'    blnCBC = False
'    blnDIFF = False
'    blnRETI = False
'    strPART = ""
    
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 파트찾기(CBC,DIFF,RET)
          SQL = "Select DISTINCT RSLTCHANNEL " & vbCr
    SQL = SQL & "  From EQPMASTER " & vbCr
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "'" & vbCr
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")" & vbCr
    SQL = SQL & " Order By RSLTCHANNEL "
        
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' "001002003004005009"
            strExamCode = strExamCode & Right("000" & Trim(AdoRs_Local.Fields("RSLTCHANNEL").Value & ""), 3)
            'strExamCode = strExamCode & Format(Trim(AdoRs_Local.Fields("RSLTCHANNEL").Value & ""), "000")
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA2120 = strExamCode
    
    
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ADVIA1800()
    Dim strOutput   As String     '송신할 데이터
    
    Select Case sSndState
        Case ""
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
            '## Order 없는 경우
            If mOrder.NoOrder = True Then
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & "000"                                                   'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "2"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8) & " 1.0" & "1" & "1"                          '
                strOutput = strOutput & Space$(1) & ETX
            Else
                '1O 0101010N003498582                                            M            1.011 89M 81M 82M 90M 91M 85M106M103M104M105M 15
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & Format$(mOrder.SendCnt, "000")                          'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "0"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8)                                               '
                strOutput = strOutput & " 1.0"                                                  'Dilution coefficient(4)
                If mOrder.SPCCD = "1" Then                                                      'Sample classification(1:blood serum, 2:urine)
                    strOutput = strOutput & "1"
                Else
                    strOutput = strOutput & "2"
                End If
                strOutput = strOutput & "1"                                                     'Container classification
                strOutput = strOutput & mOrder.Order & Space$(1) & ETX
                
            End If
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
            
            intFrameNo = intFrameNo + 1
            
        Case "E"  '## 처음 Packet 전송
            iOrderFlag = 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"  '## Packet 전송
            iOrderFlag = iOrderFlag + 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"  '## EOT
            'strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    If intFrameNo = 8 Then
        intFrameNo = 0
    End If
    
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ACLELITE()
    Dim strOutput As String     '송신할 데이터
    Dim strDate   As String
    
    Select Case intSndPhase
        Case 1  '## Header
            strDate = Format$(Now, "YYYYMMDDHHMMSS")
            strOutput = intFrameNo & "H|\^&||||||||ACL9000||P|1|" & strDate & vbCr & ETX
            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                intSndPhase = 4
            Else
                intSndPhase = 2
            End If
        
            intFrameNo = intFrameNo + 1
        
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||" & mOrder.PatId & "|||||" & mOrder.Sex & "||||||||||||||||||||||||||" & vbCr & ETX
            strOutput = intFrameNo & "P|1||" & mOrder.PID & "|||||||||||||||||||||||||||||||" & vbCr & ETX
            
            intSndPhase = 3
            
            intFrameNo = intFrameNo + 1
            
        Case 3  '## Order
            With mOrder
                'strOutput = intFrameNo & "O|" & CStr(.SendCnt + 1) & "|" & .BarNo & "||" & .Items(.SendCnt + 1) & "|" & .StatFg & "||||||||||||||||||||0||||||" & vbCr & ETX
                 strOutput = intFrameNo & "O|" & CStr(.SendCnt + 1) & "|" & .BarNo & "||" & .Items(.SendCnt + 1) & "|||||||||||||||||||||0||||||" & vbCr & ETX  '  POMIS
                 'strOutput = intFrameNo & "O|1|" & .BarNo & "||" & .Items(.SendCnt + 1) & "|||||||||||||||||||||0||||||" & vbCr & ETX  '  POMIS
                .SendCnt = .SendCnt + 1
                
                If .Count = .SendCnt Then
                    intSndPhase = 4
                Else
                    intSndPhase = 3
                End If
            End With
            
            intFrameNo = intFrameNo + 1
            
        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1
            
        Case 5  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ADVIA2120()
    Dim strOutput   As String     '송신할 데이터
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
    
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    '## Order 없는 경우
    If mOrder.NoOrder = True Then
        strOutput = msMT & "N R " & Format(mOrder.BarNo, "00000000000000") & vbCr & vbLf
    Else
        '해당 Work Order가 있는 경우
        'strOutput = msMT & "Y     " & Format(mOrder.BarNo, "00000000000000") & Space(42)
        strOutput = msMT & "Y" & Space(3) & "A" & Space(1) & Format(mOrder.BarNo, "00000000000000") & Space(42)
        strOutput = strOutput & Space(58)
        strOutput = strOutput & Space(14) & vbCr & vbLf
        strOutput = strOutput & mOrder.Order
        strOutput = strOutput & vbCr & vbLf
    End If
        
    strOutput = CheckSum_ADVIA2120(strOutput)
    
    'Delay 2 sec --> 오더쿼리시간 고려하여 Delay 1 sec
    Call Sleep(1000)
    
    comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput


End Sub

Private Sub SerialRcvData_ADVIA2120()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strKind         As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim sBC$
    Dim iStartPos       As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim strUseRes       As String
    Dim intPosS         As Integer
    
'On Error GoTo RST

    With frmMain
        strRcvBuf = RcvBuffer
        
'        strRcvBuf = "OR 00000009855100 020-08           09/01/17 16:23:19   "
'        strRcvBuf = "  1 3.17   2 2.19   3  5.7   4 15.1   5 68.8   6 26.1   7 37.9  51 38.9   8 16.8   9 1.70  10   58  11 10.1  20 41.9  21 35.8  22 14.7  23  1.3  24  0.9  25  5.3  50 2.74 "
'        strBarno = "9855100"
        
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
    
        sBC = Mid(strRcvBuf, 2, 1)
    
        Select Case sBC
            Case "S"
                Call TransferToken
                
            Case "Q"    '## Request Information
                strBarno = Trim$(Mid$(strRcvBuf, 4, 14))
                
                If IsNumeric(strBarno) Then
                    strBarno = Val(strBarno)
                End If
                
                sRcvState = "Q"
                
                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                End With
                
                Call GetOrder_ADVIA2120(strBarno, gHOSP.RSTTYPE)
                
                Call SendOrder_ADVIA2120
                
                .spdQcResult.MaxRows = 0
            Case "R"
                sRcvState = "R"
                
                strBarno = Mid(strRcvBuf, 4, 14)
                strRackNo = Mid(strRcvBuf, 19, 3)
                strTubePos = Mid(strRcvBuf, 23, 2)
                                    
                If IsNumeric(strBarno) Then
                    strBarno = Val(strBarno)
                End If
                                                    
                strKind = strQCFlag("HEMO", strBarno)
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Kind = strKind
                    .Rerun = ""
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    End If
                End With
                
                .spdQcResult.MaxRows = 0
                
                'iStartPos : Result 시작 위치
                iStartPos = InStr(strRcvBuf, vbCr)
                If iStartPos = 0 Then
                    Exit Sub
                End If
                
                iStartPos = iStartPos + 2
                
                For i = 1 To mc_iMaxCnt
                    strTemp1 = Mid(strRcvBuf, iStartPos + 9 * (i - 1), 1)
                    
                    If strTemp1 = vbCr Then Exit For
                
                    ' 특수문자 제거
                    For j = 1 To Len(strRcvBuf)
                        intPosS = InStr(strRcvBuf, "|")
                        If intPosS > 0 Then
                            strRcvBuf = Replace(strRcvBuf, "|" & Mid(Mid(strRcvBuf, intPosS + 1), 1, InStr(Mid(strRcvBuf, intPosS + 1), "|")), "")
                        Else
                            Exit For
                        End If
                    Next
                                                
                    strIntBase = CStr(Val(Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1), 3))))
                    strResult = Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1) + 3, 5))
                    strFlag = Trim(Mid(strRcvBuf, iStartPos + 9 * (i - 1) + 3 + 5, 1))
                    
                    If Left(strResult, 1) = "." Then
                        strResult = "0" & strResult
                    End If
                
                    If strIntBase <> "" And strResult <> "" And IsNumeric(strResult) Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '-- 소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strUseRes <> "" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                
'                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
'                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
'                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
'                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
'                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
'                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" And strQCAnalyte <> "" Then
                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult, strQCMethod)
                                    'Call SendBioRadQC(strQCData)
                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                Next
                
                '-- BIORAD QC 저장
                If mResult.Kind = "QC" Then
                    If .spdQcResult.MaxRows > 0 Then
                        strQCData = ""
                        For i = 1 To .spdQcResult.MaxRows
                            For j = 1 To 16
                                strQCData = strQCData & Trim(GetText(.spdQcResult, i, j)) & "|"
                            Next
                            strQCData = strQCData & vbCrLf
                        Next
                        If strQCData <> "" Then
                            Call SendBioRadQC(strQCData)
                        End If
                    End If
                End If
                                
                Call TransferResultValMsg
                
                .spdResult.RowHeight(-1) = 14
            
                '## DB에 결과저장
                If .optTrans(0).Value = True And strState = "R" Then
                    Res = SaveTransData_MCC(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "Failed", gRow, colSTATE
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
    End With
'Exit Sub
'RST:
'
'    Screen.MousePointer = vbDefault
'    MsgBox " Error No. : " & Err.Number & vbCrLf & _
'            " Description : " & Err.Description & vbCrLf & _
'            " Source : " & Err.Source & vbCrLf & vbCrLf
'
End Sub

Private Sub TransferToken()
    Dim sSendBuf$
    
    '프로그램 종료 예정 --> 장비쪽이 Slave인 상태에서 종료되도록 TransferToken 안함
    If mp_bReserveEnd Then
        frmMain.comEqp.PortOpen = False
        mp_bPortOpen = False
        Exit Sub
    End If
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
   
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "S " & vbCr & vbLf
    
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    msSndPacket = sSendBuf

    Timer2120.Interval = 5000
    Timer2120.Enabled = True

End Sub


Private Sub TransferResultValMsg()
    'Result Validation Message
    
    Dim sSendBuf$
    
    If msMT = "" Then msMT = Chr(&H30)

    msMT = Chr(Asc(msMT) + 1)
        
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "Z                  0" & vbCr & vbLf
    sSendBuf = CheckSum_ADVIA2120(sSendBuf)
    
    'Delay 1.5 sec --> 결과등록시간 고려하여 Delay 1 sec
    Call Sleep(1000)
    
    comEqp.Output = sSendBuf
    SetRawData "[Tx]" & sSendBuf

End Sub

Private Sub VesMatic(asData As String)
    Dim strHeader As String
    
    Call SetSQLData("RCV", asData, "A")
    
    strHeader = Trim(mGetP(asData, 1, "="))
    
    If strHeader <> "" And IsNumeric(strHeader) Then
        Call SerialRcvData_VESCUBE
    Else
        Exit Sub
    End If
    
End Sub

Private Sub SerialRcvData_VESCUBE()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRecvData = Split(RcvBuffer, vbCr)
        For intCnt = 1 To UBound(strRecvData)
            RcvBuffer = strRecvData(intCnt)
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", RcvBuffer, "A")
            End If
            '-- 테스트용 -----------------
                
            '1 = O4ZU70QN0....  48
            '1 = 199297.......   6
            strTemp1 = mGetP(RcvBuffer, 2, "=")
            strBarno = Trim(mGetP(strTemp1, 1, "......."))
    
            If Trim(strBarno) <> "" And Len(strBarno) = 6 Then
                With mResult
                    .BarNo = strBarno
                    '.SpcPos = strTubePos & "/" & strRackNo
                    '.Seq = strSeq
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    End If
                End With
                
                If gPatOrdCd <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                        '-- 결과Row 추가
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE
    
                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        '-- BIORAD QC 저장
                        If Mid(strBarno, 1, 2) = "QC" Then
                            Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                        End If
                    
                        
                        strState = "R"
                        
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                        
                    End If
                Else
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                        strQCLab = Trim(RS_L.Fields("QCLab") & "")
                        strQCLot = Trim(RS_L.Fields("QCLot") & "")
                        strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                        strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                        strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                        strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                        strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                        strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                        '-- 결과Row 추가
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If
    
                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE
    
                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
    
                        '-- 결과 List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        '-- BIORAD QC 저장
                        If Mid(strBarno, 1, 2) = "QC" Then
                            Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                        End If
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
    
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                    End If
                    
                End If
                
            End If
                            
            .spdResult.RowHeight(-1) = 14
                
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
        Next
    End With

End Sub



Private Sub Phase_Serial_VESCUBE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    TimerVESCUBE.Interval = 5000
    TimerVESCUBE.Enabled = True
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case BufChar
            Case vbLf
            Case vbCr
                    Call VesMatic(RcvBuffer)
                    RcvBuffer = ""
            
            Case Else
                    RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_ADVIA2120()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

'    SetRawData "[Rx]" & pBuffer
    lngBufLen = Len(pBuffer)
            
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1                  ' 초기화 확인(MT 대기)
                Select Case BufChar
                    Case Chr(2)          'STX
                        RcvBuffer = ""
                        
                    Case msMT                           ' 0(&H30), Initialize Message 직후이므로
                        RcvBuffer = ""
                        Call TransferToken              'Transfer_Token 시도 후
                        intPhase = 2

                    Case Else
                        'MT이외의 경우(즉, NACK)
                        Call InitialComm
                        intPhase = 1
                        Exit Sub
                End Select

            Case 2                                      ' Token Tranfer에 대한 MT 대기
                Select Case BufChar
                    Case Chr(21)         'NAK
                        Call InitialComm
                        intPhase = 1
                        
                    Case msMT
                        intPhase = 3

                    Case Chr(Asc(msMT) - 1)
                        msMT = Chr(Asc(msMT) - 1)
                        If Asc(msMT) < 0 Then
                            msMT = Chr(&H30)
                        End If
                        Call TransferToken              'Transfer_Token 시도 후
                        intPhase = 2

                    Case Else
                        Call TransferToken
                        intPhase = 2
                End Select

            Case 3                                      ' CheckSum STX인 경우의 오류 방지를 위해, Phase 3 과 4 분리
                Select Case Asc(BufChar)
                    Case 2          'STX
                        RcvBuffer = ""
                        intPhase = 4

                    Case Else
                        intPhase = 3

                End Select

            Case 4                                      'DataEdit(장비쪽에서 보내는 S, Q, R 메세지에 대한) 대기,
                Select Case Asc(BufChar)
                    Case 3            ' ETX
                        msMT = Left(RcvBuffer, 1)
                        Call Sleep(25)  'Delay Time 0.025 sec

                        comEqp.Output = msMT
                        SetRawData "[Tx]" & msMT
                        Call SerialRcvData_ADVIA2120
                        intPhase = 3

                    Case Else
                        RcvBuffer = RcvBuffer & BufChar

                End Select
        End Select
    Next i
    
    
End Sub

''-----------------------------------------------------------------------------'
''   기능 : 오더정보 전송
''-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'    Dim strOutput   As String     '송신할 데이터
'
''1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20161209210918
''6B
''2P|1|03217192|||Jo^ Yu Jeong^^|||U
''65
''3O|1|03217192||^^^wrCRP\^^^AMYLAS|R||||||||||1
''48
''4L|1|N
''07
'
'    Select Case intSndPhase
'        Case 1  '## Header
'            strOutput = intFrameNo & "H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|" & Format(Now, "yyyymmddhhmmss") & "|" & vbCr & ETX
'            intSndPhase = 2
'            intFrameNo = intFrameNo + 1
'
'        Case 2  '## Patient
'            strOutput = intFrameNo & "P|1|" & mOrder.BarNo & "|||" & frmMain.Han2Eng.HanToEng(mOrder.PName) & "||||" & vbCr & ETX
'            intSndPhase = 3
'            intFrameNo = intFrameNo + 1
'
'        Case 3  '## Order
'            If mOrder.NoOrder = True Then
'                '## 접수정보가 없을경우
'                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.SPCCD
'                strOutput = intFrameNo & strOutput & vbCr & ETX
'                intSndPhase = 4
'
'            Else
'                '## 최초 보낼때
'                If mOrder.IsSending = False Then
'                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||1"
'
'                    If Len(strOutput) > 230 Then
'                        mOrder.IsSending = True
'                        mOrder.Order = Mid$(strOutput, 231)
'                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                        intSndPhase = 3
'                    Else
'                        strOutput = intFrameNo & strOutput & vbCr & ETX
'                        intSndPhase = 4
'                    End If
'                '## 남은 문자열이 있을때
'                Else
'                    strOutput = mOrder.Order
'                    If Len(strOutput) > 230 Then
'                        mOrder.Order = Mid$(strOutput, 231)
'                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                        intSndPhase = 3
'                    Else
'                        mOrder.IsSending = False
'                        strOutput = intFrameNo & strOutput & vbCr & ETX
'                        intSndPhase = 4
'                    End If
'                End If
'            End If
'            intFrameNo = intFrameNo + 1
'
'        Case 4  '## Termianator
'            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
'            intSndPhase = 5
'            intFrameNo = intFrameNo + 1
'
'        Case 5  '## EOT
'            strState = ""
'            frmMain.comEqp.Output = EOT
'            SetRawData "[Tx]" & EOT
'            intFrameNo = 1
'
'            Exit Sub
'    End Select
'
'    If intFrameNo = 8 Then
'        intFrameNo = 0
'    End If
'
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput
'
'End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_VERSACELL()
    Dim strOutput   As String     '송신할 데이터

'1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20161209210918
'6B
'2P|1|03217192|||Jo^ Yu Jeong^^|||U
'65
'3O|1|03217192||^^^wrCRP\^^^AMYLAS|R||||||||||1
'48
'4L|1|N
'07

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|" & Format(Now, "yyyymmddhhmmss") & "|" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|" & mOrder.BarNo & "|||" & frmMain.Han2Eng.HanToEng(mOrder.PName) & "||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1

        Case 3  '## Order
            If mOrder.NoOrder = True Then
                '## 접수정보가 없을경우
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.SPCCD
                strOutput = intFrameNo & strOutput & vbCr & ETX
                intSndPhase = 4

            Else
                '## 최초 보낼때
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||1"

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
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub

Private Sub Phase_Serial_VERSACELL()
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_VERSACELL
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_VERSACELL
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub SerialRcvData_UrinscanPro()
    Dim RS_L            As ADODB.Recordset
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
    Dim strFunction     As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim strQCTemp       As String
    Dim Pos As Integer
    
    
    
    With frmMain
        Pos = InStr(RcvBuffer, "ID_NO")
        If Pos > 0 Then
            With mResult
                .BarNo = strBarno
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                .SpcPos = strSeq
            End With
                    
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
            RcvBuffer = Replace(RcvBuffer, vbLf, "")
            strRecvData = Split(RcvBuffer, vbCr)
            
            
            For intCnt = 4 To UBound(strRecvData)
                strRcvBuf = strRecvData(intCnt)
                
                '-- 테스트용 -----------------
                'If .fraCommTest.Visible = False Then
                '    Call SetSQLData("RCV", strRcvBuf, "A")
                'End If
                '-- 테스트용 -----------------
                
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
                        
    
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Next
                
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_GINUS(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
        End If
    End With

End Sub

Private Sub HTMLRcvData_ALLEREi()
    Dim RS_L            As ADODB.Recordset
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
    Dim strFunction     As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strFunc         As String
    Dim i               As Integer
    Dim j               As Integer
    Dim strQCTemp       As String
    Dim Pos As Integer
    Dim strFluA As String
    Dim strFluB As String
    Dim varTmp As Variant
    
    With frmMain
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            'If .fraCommTest.Visible = False Then
            '    Call SetSQLData("RCV", strRcvBuf, "A")
            'End If
            '-- 테스트용 -----------------
            
            strBarno = Trim(mGetP(strRcvBuf, 1, "|"))
            
            With mResult
                .BarNo = strBarno
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
            End With
                    
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            
            If gRow <= 0 Then
                Exit Sub
            End If
            
            strState = "O"
            
            '-- 오른쪽 결과화면 초기화
            spdResult.MaxRows = 0
            
            '0 : Neg ,1 : Pos, -1 : Neg
            For i = 1 To 3
                '## 장비기준 검사명, 결과, Abnormal Flag
                If i = 1 Then
                    strIntBase = "FluA"
                    strResult = Trim(mGetP(strRcvBuf, 2, "|"))
                    If strResult = "0" Or strResult = "-1" Then
                        strFluA = "NEGATIVE"
                        strResult = "NEGATIVE"
                    ElseIf strResult = "1" Then
                        strFluA = "POSITIVE"
                        strResult = "POSITIVE"
                    End If
                ElseIf i = 2 Then
                    strIntBase = "FluB"
                    strResult = Trim(mGetP(strRcvBuf, 3, "|"))
                    If strResult = "0" Or strResult = "-1" Then
                        strFluB = "NEGATIVE"
                        strResult = "NEGATIVE"
                    ElseIf strResult = "1" Then
                        strFluB = "POSITIVE"
                        strResult = "POSITIVE"
                    End If
'                Else
'                    strIntBase = "Flu"
'
'                    If strFluA = "NEGATIVE" Then
'                        If strFluB = "NEGATIVE" Then
'                            strResult = "NEGATIVE"
'                        Else
'                            strResult = "POSITIVE(type A)"
'                        End If
'                    ElseIf strFluA = "POSITIVE" Then
'                        If strFluB = "NEGATIVE" Then
'                            strResult = "POSITIVE(type B)"
'                        Else
'                            strResult = "POSITIVE(type A & B)"
'                        End If
'                    End If
                End If

                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
                            lsOrderCode = ""
                            
                            '반환장비검사코드 찾기
                            If gPatTstCd <> "" Then
                                varTmp = Split(gPatTstCd, "|")
                                For j = 0 To UBound(varTmp)
                                    If Trim(mGetP(varTmp(i), 1, ",")) = lsTestCode Then
                                        lsOrderCode = Trim(mGetP(varTmp(i), 2, ","))
                                        Exit For
                                    End If
                                Next
                            End If
                            
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Next
            
            '## DB에 결과저장
            'If .optTrans(0).Value = True And strState = "R" Then
            If .optTrans(0).Value = True Then
                Res = SaveTransData_BDJS(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
        Next
    End With
    
End Sub


Private Sub Phase_Serial_OSMOPRO()
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            'Call SendOrder_VERSACELL
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_OSMOPRO
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
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
Private Sub GetOrder_Versacell(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
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
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_VERSACELL(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
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
Private Sub GetOrder_ACLELITE(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
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
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ACLELITE(gHOSP.MACHCD, pBarno, intRow)

        Call SetSQLData("strItems ", strItems)
        
        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_VERSACELL(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_VERSACELL = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_VERSACELL = Mid(strExamCode, 2)
    
End Function

Private Function GetEquipExamCode_ACLELITE(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String
    Dim strTemp     As String
    Dim strIntBase  As String

    GetEquipExamCode_ACLELITE = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Call SetSQLData("채널찾기 ", SQL)
    
    strExamCode = ""
    mOrder.Count = 0
    Erase mOrder.Items
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
                        'If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
            strIntBase = Mid(Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 1, 4)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & "^^^" & strIntBase
                mOrder.Count = mOrder.Count + 1
                ReDim Preserve mOrder.Items(mOrder.Count)
                mOrder.Items(mOrder.Count) = "^^^" & strIntBase
                strTemp = strIntBase
            End If
                        'End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ACLELITE = strExamCode
    
End Function

Private Sub SerialRcvData_VERSACELL()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                    '2P|1|Multi QC Lv.1|||45731     /||||||||
                    '2P|1|Multi QC Lv3|||45733     /||||||||
                    '2P|1|AMMONIA QC Lv.1|||54181     /||||||||
                    '2P|1|AMMONIA QC Lv.3|||54183     /||||||||
                    
                    If InStr(mGetP(strRcvBuf, 3, "|"), "QC") > 0 Then
                        mResult.Kind = "QC"
                        .spdQcResult.MaxRows = 0
                    Else
                        mResult.Kind = ""
                    End If

                Case "Q"    '## Request Information
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .Seq = mGetP(strTemp1, 3, "^")
                        .RackNo = mGetP(strTemp1, 4, "^")
                        .TubePos = mGetP(strTemp1, 5, "^")
                    End With
                    
                    Call GetOrder_ACLELITE(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
'                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                            
                            Call CalProcess(spdOrder, spdResult, lsTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                        
                    mResult.EqpCd = ""
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    If strBarno = "" Then Exit Sub
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = UCase(Mid(mGetP(strRcvBuf, 5, "|"), 1, 1))

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                
                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                    Call SendBioRadQC(strQCData)
                                End If
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                            
                            Call CalProcess(spdOrder, spdResult, lsTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    
'                Case "L"
'                    '-- BIORAD QC 저장
'                    If mResult.Kind = "QC" Then
'                        If .spdQcResult.MaxRows > 0 Then
'                            strQCData = ""
'                            For i = 1 To .spdQcResult.MaxRows
'                                For j = 1 To 16
'                                    strQCData = strQCData & Trim(GetText(.spdQcResult, i, j)) & "|"
'                                Next
'                                strQCData = strQCData & vbCrLf
'                            Next
'                            If strQCData <> "" Then
'                                Call SendBioRadQC(strQCData)
'                            End If
'                        End If
'                    End If
            End Select
        Next
    End With

End Sub

Private Sub SerialRcvData_ACLELITE()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    Dim strPTsec        As String
    Dim strPTINR        As String
    Dim strPTPer        As String
    
    Dim strQCFileName   As String
    
    With frmMain
        strQCFileName = Format(Now, "yyyymmddhhmmss")
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- Easy QC 용 -----------------
            Call SetLogData(strQCFileName, strRcvBuf, "A")
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .Seq = mGetP(strTemp1, 3, "^")
                        .RackNo = mGetP(strTemp1, 4, "^")
                        .TubePos = mGetP(strTemp1, 5, "^")
                        .SendCnt = 0
                        .Count = 0
                    End With
                    
                    Call GetOrder_ACLELITE(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                    
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                Case "R"
'4R|1|^^^0013|10.500   |s||||F||||20171206235634|
'5R|2|^^^0013|0.991    |R||||F||||20171206235634|
'6R|3|^^^0013|0.990    |INR||||F||||20171206235634|
'7R|4|^^^0013|105.058  |%||||F||||20171206235634|


                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = UCase(Mid(mGetP(strRcvBuf, 5, "|"), 1, 1))
                    
                    strIntBase = strIntBase & strFlag
                    
                    If strIntBase = "0013S" Then
                        strPTsec = Format(strResult, "##.0") & "s"
                        'strResult = ""
                    End If

                    If strIntBase = "0013I" Then
                        strPTINR = Format(strResult, "##.#0") & "R"
                        'strResult = ""
                    End If

                    If strIntBase = "0013%" Then
                        strPTPer = Format(strResult, "#0") & "%"
                        strResult = strPTsec & " " & strPTPer & " " & strPTINR
                        strPTsec = ""
                        strPTINR = ""
                    End If

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(lsTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                                            
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            'SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & " WHERE RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                'strResult = getCRRValue(lsTestCode, strResult)
                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                                                
                                
                                'If strState <> "R" Then
                                    strState = ""
                                'End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                
                
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_JWINFO(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                            
                            'Call CalProcess(spdOrder, spdResult, lsTestCode)
                            
                        End If
                        strState = ""
                        
                    End If
                    

            End Select
        Next
    End With

End Sub

Public Sub SerialRcvData_OSMOPRO()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    Dim strKind         As String
    
    Dim i               As Integer
    Dim j               As Integer
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                Case "O"
                    'O|1|3MA005||^^^||20161110082002|20161110082002||0.000|||||20161110082002||||||||20161110082002|||<CR><ETX
                    
                    strBarno = mGetP(strRcvBuf, 3, "|")
                    
                    If UCase(strBarno) = "S1" Or UCase(strBarno) = "S3" Or UCase(strBarno) = "U1" Or UCase(strBarno) = "U2" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        .Kind = strKind
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    
                                        
                Case "R"
                    'R|1|^^^OSMO|51|mOsm/Kg H2O||N|N|F||OperatorID|20161027142723|| 17010095A<CR><ETX>
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                    
                                    Call SendBioRadQC(strQCData)
                                    
                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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

End Sub

Public Sub TCPRcvData_ALEREi()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    'Dim strOldBarno        As String   '수신한 바코드번호
    Dim strBarno        As String   '수신한 바코드번호
    Dim strQCYN         As String   '수신한 바코드번호
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
    Dim strAspect       As String
    Dim strEqpNm        As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim strINTRResult   As String
    Dim strKind         As String
    
    Dim i               As Integer
    Dim j               As Integer
    Dim varTmp          As Variant
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "O"
                    'O|1||1739a3ff-b97d-4aa8-8e3b-0bb3df39d315|^^^Influenza A &E& B|||||||||||Vtm^^M094440||||||||||F
                    strQCYN = mGetP(mGetP(strRcvBuf, 16, "|"), 1, "^")
                    If strQCYN = "QC" Then
                        Exit Sub
                    End If
                Case "P"
                    'P|1||specimen(a1)
                    strBarno = mGetP(strRcvBuf, 4, "|")
                    If Trim(strBarno) = "" Then
                        Exit Sub
                    End If
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    
                                        
                Case "R"
                    'R|1|^^^Flu A^M|Positive|||A||F||admin|20180125192309|20180125193311|54D9A118
                    'R|2|^^^Flu B^M|Negative|||N||F|||||54D9A118
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strResult = mGetP(strRcvBuf, 4, "|")
                    
'                    If UCase(strResult) = "NEGATIVE" Then
'                        strResult = "음성"
'                    Else
'                        strResult = "양성"
'                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
                                
                                'gPatTstCd = "X3B000087_003,IA|X3B000087_004,IB|"
                                varTmp = Split(gPatTstCd, "|") 'X3B000087_003,IA|X3B000087_004,IB|
                                
                                For i = 0 To UBound(varTmp)
                                    If lsTestCode = mGetP(varTmp(i), 1, ",") Then
                                        lsOrderCode = mGetP(varTmp(i), 2, ",")
                                        Exit For
                                    End If
                                Next
                                '
        
                                'X3B000087_003,IA|X3B000087_004,IB|
                                
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                                                                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'
'                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'
'                                End If
                                
                                If strState <> "R" Then
                                    strState = "R"
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True Then
                        Res = SaveTransData_BDJS(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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

End Sub



'''Public Sub SerialRcvData_CT500()
'''    Dim RS_L            As ADODB.Recordset
'''    Dim strRcvBuf       As String   '수신한 Data
'''    Dim strType         As String   '수신한 Record Type
'''    Dim strBarno        As String   '수신한 바코드번호
'''    Dim strSeq          As String   '수신한 Sequence
'''    Dim strRackNo       As String   '수신한 Rack Or Disk No
'''    Dim strTubePos      As String   '수신한 Tube Position
'''    Dim strIntBase      As String   '수신한 장비기준 검사명
'''    Dim strMachResult   As String   '수신한 장비결과
'''    Dim strResult       As String   '수신한 결과(정성)
'''    Dim strIntResult    As String   '수신한 결과(정량)
'''    Dim varResult       As Variant
'''    Dim strQCResult     As String   '수신한 결과(QC)
'''    Dim strFlag         As String   '수신한 Abnormal Flag
'''    Dim strComm         As String   '수신한 Comment
'''
'''    Dim lsOrderCode     As String   '처방코드
'''    Dim lsTestCode      As String   '검사코드
'''    Dim lsTestName      As String   '검사명
'''    Dim lsSeqNo         As String   '로컬DB 검사Seq
'''
'''    Dim lsRstRow        As String   '결과스프레드 현재 Row
'''    Dim intCnt          As Integer  '통신 Frame 갯수
'''    Dim intCol          As Integer  '결과컬럼 갯수
'''    Dim strJudge        As String   '결과판정
'''    Dim Res             As Integer
'''
'''    Dim strTmp          As String
'''    Dim strOldBarno     As String
'''    Dim strQCData       As String
'''    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'''
'''    With frmMain
'''        strRcvBuf = RcvBuffer
'''        strRcvBuf = Replace(strRcvBuf, vbLf, "")
'''
''''#4-723      17-08-28
''''ID = 3495464
''''Color: STRAW
''''Clarity:
''''GLU NEGATIVE
''''BIL NEGATIVE
''''KET NEGATIVE
''''SG 1.025
''''BLO NEGATIVE
''''pH 6#
''''PRO NEGATIVE
''''URO      0.2 E.U./dL
''''NIT NEGATIVE
''''LEU NEGATIVE
''''
'''
'''
'''        '-- 테스트용 -----------------
'''        If .fraCommTest.Visible = False Then
'''            Call SetSQLData("RCV", strRcvBuf, "A")
'''        End If
'''        '-- 테스트용 -----------------
'''
'''        If Mid(strRcvBuf, 1, 3) = "ID=" Then
'''            miLineNo = 1
'''            mColor = False
'''            'strBarno = Trim(Mid(strRcvBuf, 5, 12))
'''            strBarno = Trim(Mid(strRcvBuf, 4))
'''            mResult.BarNo = strBarno
'''            If strBarno = "1" Or strBarno = "2" Then
'''                mResult.Kind = "QC"
'''            End If
'''
'''            With mResult
'''                .BarNo = strBarno
'''                .SpcPos = strSeq
'''                .Seq = strSeq
'''                .RackNo = strRackNo
'''                .TubePos = strTubePos
'''                If strOldBarno <> strBarno Then
'''                    strOldBarno = strBarno
'''                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'''                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'''                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''                End If
'''            End With
'''
'''            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'''
'''            If gRow <= 0 Then
'''                Exit Sub
'''            End If
'''        Else
'''            If miLineNo = 1 Then
'''                strIntBase = Trim(mGetP(strRcvBuf, 1, Space$(1)))
'''                If Right(strIntBase, 1) = "*" Then
'''                    strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1)
'''                End If
'''                strResult = Trim(mGetP(strRcvBuf, 2, Space$(1)))
'''                If strResult = "" Then
'''                    If Len(strIntBase) = 3 Then
'''                        strResult = Trim(Mid(strRcvBuf, 8))
'''                    Else
'''                        strResult = Trim(Mid(strRcvBuf, 9))
'''                    End If
'''                End If
'''                strResult = Replace(strResult, "E.U./dL", "")
'''                strResult = Trim(strResult)
'''
'''
'''                If strIntBase = "Color:" Then
'''                    mColor = True
'''                End If
'''
'''                '--QC
'''                If Len(mResult.BarNo) <= 5 Then
'''                    strResult = Replace(strResult, "<", "")
'''                    strResult = Replace(strResult, ">", "")
'''                    strResult = Replace(strResult, "=", "")
'''                End If
'''
'''RST:
'''                If strIntBase <> "" And strResult <> "" Then
'''                    If gPatOrdCd <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp " & vbCr
'''                        SQL = SQL & "  FROM EQPMASTER" & vbCr
'''                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'''                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'''                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'''
'''                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'''                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'''                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'''                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
'''                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'''                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
'''
'''                            '-- 결과Row 추가
'''                            lsRstRow = .spdResult.DataRowCnt + 1
'''                            If .spdResult.MaxRows < lsRstRow Then
'''                                .spdResult.MaxRows = lsRstRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            strMachResult = strResult
'''                            If strQCTemp = "1" Then
'''                                strResult = SetResult(strResult, strIntBase)
'''                            End If
'''                            strJudge = SetJudge(strResult, strIntBase)
'''
'''                            '진행상태 표시("결과")
'''                            SetText .spdOrder, "결과", gRow, colSTATE
'''
'''                            '결과값 표시
'''                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'''                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'''                                    SetText .spdOrder, strResult, gRow, intCol
'''                                    Exit For
'''                                End If
'''                            Next
'''
'''                            '-- 결과 List
'''                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'''                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'''                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'''                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'''                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'''                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'''                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'''                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'''                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'''
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, lsRstRow, "1", ""
'''
'''                            strState = "R"
'''
'''                            '-- BIORAD QC 저장
''''                            If mResult.Kind = "QC" Then
''''                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
''''
''''                                Call SendBioRadQC(strQCData)
''''                            End If
'''
'''                            '-- 결과Count
'''                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'''                                SetText .spdOrder, "1", gRow, colRCNT
'''                            Else
'''                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'''                            End If
'''
'''                        End If
'''                    Else
'''                        SQL = ""
'''                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp,QCAnalyte " & vbCr
'''                        SQL = SQL & "  FROM EQPMASTER" & vbCr
'''                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'''                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'''
'''                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'''                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'''                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'''                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'''                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'''                            strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
'''                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'''
'''                            '-- 결과Row 추가
'''                            lsRstRow = .spdResult.DataRowCnt + 1
'''                            If .spdResult.MaxRows < lsRstRow Then
'''                                .spdResult.MaxRows = lsRstRow
'''                            End If
'''
'''                            '소수점 처리, 결과 형태 처리
'''                            strMachResult = strResult
'''                            If strQCTemp = "1" Then
'''                                strResult = SetResult(strResult, strIntBase)
'''                            End If
'''                            strJudge = SetJudge(strResult, strIntBase)
'''
'''                            '진행상태 표시("결과")
'''                            SetText .spdOrder, "결과", gRow, colSTATE
'''
'''                            '결과값 표시
'''                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'''                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'''                                    SetText .spdOrder, strResult, gRow, intCol
'''                                    Exit For
'''                                End If
'''                            Next
'''
'''                            '-- 결과 List
'''                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'''                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'''                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'''                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'''                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'''                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'''                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'''                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'''                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'''
'''                            '-- 로컬 저장
'''                            SetLocalDB gRow, lsRstRow, "1", ""
'''
'''                            If strState <> "R" Then
'''                                strState = ""
'''                            End If
'''
'''                            '-- BIORAD QC 저장
'''                            If mResult.Kind = "QC" Then
'''                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, mResult.BarNo, strQCAnalyte, strResult)
'''
'''                                Call SendBioRadQC(strQCData)
'''                            End If
'''
'''                            '-- 결과Count
'''                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
'''                                SetText .spdOrder, "1", gRow, colRCNT
'''                            Else
'''                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'''                            End If
'''                        End If
'''
'''                    End If
'''
'''                End If
'''
'''                .spdResult.RowHeight(-1) = 14
'''
'''                '## DB에 결과저장
'''                If .optTrans(0).Value = True And strState = "R" Then
'''                    Res = SaveTransData_MCC(gRow)
'''
'''                    If Res = -1 Then
'''                        '-- 저장 실패
'''                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'''                        SetText .spdOrder, "Failed", gRow, colSTATE
'''                    Else
'''                        '-- 저장 성공
'''                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'''                        SetText .spdOrder, "저장완료", gRow, colSTATE
'''                        SetText .spdOrder, "0", gRow, colCHECKBOX
'''
'''                              SQL = "Update PATRESULT Set " & vbCrLf
'''                        SQL = SQL & " sendflag = '2' " & vbCrLf
'''                        SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'''                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'''                        SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'''                        SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'''
'''                        If DBExec(AdoCn_Local, SQL) Then
'''                            '-- 성공
'''                        End If
'''                    End If
'''                    strState = ""
'''                End If
'''
'''                If mColor = False And strIntBase = "LEU" Then
'''                    strIntBase = "Color:"
'''                    strResult = "YELLOW"
'''                    GoTo RST
'''                End If
'''            End If
'''        End If
'''    End With
'''
'''End Sub

Public Sub SerialRcvData_CT500()
    Dim RS_L            As ADODB.Recordset
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
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
'    With frmMain
'        strRcvBuf = RcvBuffer
'        strRcvBuf = Replace(strRcvBuf, vbLf, "")
        
'#4-723      17-08-28
'ID = 3495464
'Color: STRAW
'Clarity:
'GLU NEGATIVE
'BIL NEGATIVE
'KET NEGATIVE
'SG 1.025
'BLO NEGATIVE
'pH 6#
'PRO NEGATIVE
'URO      0.2 E.U./dL
'NIT NEGATIVE
'LEU NEGATIVE
'

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            If Mid(strRcvBuf, 1, 3) = "ID=" Then
                strBarno = Trim(Mid(strRcvBuf, 4))
                mResult.BarNo = strBarno
                If strBarno = "1" Or strBarno = "2" Then
                    mResult.Kind = "QC"
                End If
                
                With mResult
                    .BarNo = strBarno
                    .SpcPos = strSeq
                    .Seq = strSeq
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                If gRow <= 0 Then
                    Exit Sub
                End If
            Else
                If intCnt > 2 Then
                    strIntBase = Trim(mGetP(strRcvBuf, 1, Space$(1)))
                    If Right(strIntBase, 1) = "*" Then
                        strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1)
                    End If
                    strResult = Trim(mGetP(strRcvBuf, 2, Space$(1)))
                    If strResult = "" Then
                        If Len(strIntBase) = 3 Then
                            strResult = Trim(Mid(strRcvBuf, 8))
                        Else
                            strResult = Trim(Mid(strRcvBuf, 9))
                        End If
                    End If
                    strResult = Replace(strResult, "E.U./dL", "")
                    strResult = Trim(strResult)
                    
                    
                    If strIntBase = "Color:" Then
                        mColor = True
                    End If
                        
                    '--QC
                    If Len(mResult.BarNo) <= 5 Then
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
RST:
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP"))
                                
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
            
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
            
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                strState = "R"
                                
                                '-- BIORAD QC 저장
    '                            If mResult.Kind = "QC" Then
    '                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
    '
    '                                Call SendBioRadQC(strQCData)
    '                            End If
                        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH, QCTemp,QCAnalyte " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
            
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
            
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
            
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
            
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
            
                                '-- BIORAD QC 저장
                                If mResult.Kind = "QC" Then
                                    strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, mResult.BarNo, strQCAnalyte, strResult)
                                    
                                    Call SendBioRadQC(strQCData)
                                End If
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                            
                        End If
                        
                    End If
                                
                    .spdResult.RowHeight(-1) = 14
                                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                    
                    If mColor = False And strIntBase = "LEU" Then
                        strIntBase = "Color:"
                        strResult = "YELLOW"
                        GoTo RST
                    End If
                End If
            End If
        Next
    End With

End Sub




Private Sub Phase_Serial_CT500()
     Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                RcvBuffer = ""
'
'                miLineNo = 0
'            Case vbCr
'                Call SerialRcvData_CT500
'
'                miLineNo = 1
'
'                RcvBuffer = ""
'
'            Case ETX
'                RcvBuffer = ""
'                miLineNo = 0
'
'            Case Else
'                RcvBuffer = RcvBuffer & BufChar
'        End Select
'    Next i

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                'If intBufCnt = 0 Then
                    intBufCnt = 1
                    Erase strRecvData
                    ReDim Preserve strRecvData(intBufCnt)
                'Else
                '    intBufCnt = intBufCnt + 1
                '    ReDim Preserve strRecvData(intBufCnt)
                'End If
            Case vbCr
                intBufCnt = intBufCnt + 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbLf
            
            Case ETX
                Call SerialRcvData_CT500
                Erase strRecvData
                intBufCnt = 0
                
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

End Sub

Private Sub Phase_Serial_RAPIDLAB348()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1      '## STX 대기
                Select Case BufChar
                    Case STX
                        intPhase = 2
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                
                End Select
            Case 2      '## ETX 대기
                Select Case BufChar
                    Case ETX
                        intPhase = 3
                    Case Else
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End Select
            Case 3      '## EOT 대기
                Select Case BufChar
                    Case EOT
                        Call SerialRcvData_RAPIDLAB348
                        intPhase = 1
                End Select
        End Select
    Next i
    
End Sub


Private Sub SerialRcvData_RAPIDLAB348()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    Dim strQCChannel    As String
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strIDRecord     As String   '수신한 Identifyer Record
    Dim strWorkNo       As String   '수신한 WorkNo
    Dim AssayNm         As String
    
    Dim Pos1            As Long
    Dim Pos2            As Long
    Dim x1              As Long
    Dim x2              As Long
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strctHb     As String
    Dim strO2SAT    As String
    Dim strPO2      As String

    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strIDRecord = Trim$(mGetP(strRcvBuf, 1, FS))
            
            If strIDRecord = "SMP_NEW_DATA" Or strIDRecord = "SMP_EDIT_DATA" Then
                '## WorkNo 조회
                Pos1 = InStr(strRcvBuf, "rSEQ")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strSeq = Format$(mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS), "#####")
                    strSeq = Val(strSeq)
                Else
                    '## NOTE: WorkNo가 전송되지 않은 에러처리
                    Exit Sub
                End If
                
                '## 바코드번호 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iPID")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strBarno = Format$(mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS), String$(9, "#"))
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Rerun = ""
                    'If strOldBarno <> strBarno Then
                        'strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    'End If
                End With
                          
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "m") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "m")
                    x2 = InStr(x1, strRcvBuf, GS)
        
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
                    
                    If strIntBase = "mPO2" Then
                        strPO2 = strResult
                    End If

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strQCTemp = Trim(RS_L.Fields("SEQNO") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'
'                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                Loop
                
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "c")
                    x2 = InStr(x1, strRcvBuf, GS)
            
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
            
                    If strIntBase = "ctHb(est)" Then
                        strctHb = strResult
                    End If
                    
                    If strIntBase = "cO2SAT" Then
                        strO2SAT = strResult
                    End If
                
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'                                End If
                                
                                strState = "R"
                                
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
        
                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
        
                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If
        
                                '소수점 처리, 결과 형태 처리
                                strMachResult = strResult
                                If strQCTemp = "1" Then
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)
                                
                                'CRR 적용
                                strResult = getCRRValue(lsTestCode, strResult)
                                
                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE
        
                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
        
                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                
                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""
                                
                                '-- BIORAD QC 저장
'                                If mResult.Kind = "QC" Then
'
'                                    strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                    Call SendBioRadQC(strQCData)
'
'                                End If
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
        
                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If
                        End If
                    End If
                Loop
            End If
            
            'O2CT = (1.39ctHb x O2SAT/100) + (0.00314pO2)
            strResult = ""
            If strctHb <> "" And strO2SAT <> "" And strPO2 <> "" Then
                strResult = ((1.39 * strctHb) * (strO2SAT / 100)) + (0.00314 * strPO2)
                strResult = Format(strResult, "##.00")
                strResult = Mid(strResult, 1, InStr(strResult, ".") + 1)
                strIntBase = "O2CT"
            End If
            
            If strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                SQL = SQL & "  FROM EQPMASTER" & vbCr
                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                    lsTestName = Trim(RS_L.Fields("TESTNAME"))
                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))

                    '-- 결과Row 추가
                    lsRstRow = .spdResult.DataRowCnt + 1
                    If .spdResult.MaxRows < lsRstRow Then
                        .spdResult.MaxRows = lsRstRow
                    End If

                    '소수점 처리, 결과 형태 처리
                    strMachResult = strResult
                    If strQCTemp = "1" Then
                        strResult = SetResult(strResult, strIntBase)
                    End If
                    strJudge = SetJudge(strResult, strIntBase)
                    
                    'CRR 적용
                    strResult = getCRRValue(lsTestCode, strResult)
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE

                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- 결과 List
                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    '-- BIORAD QC 저장
                    If mResult.Kind = "QC" Then
                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                        
                        Call SendBioRadQC(strQCData)
                    End If
                    
                    strState = "R"
                    
                    '-- 결과Count
                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                        SetText .spdOrder, "1", gRow, colRCNT
                    Else
                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                    End If
                End If
            End If
            .spdResult.RowHeight(-1) = 14
        
            '#########  QC Define ##########################################################
            If strIDRecord = "QC_NEW_DATA" Or strIDRecord = "QC_EDIT_DATA" Then
                
                .spdQcResult.MaxRows = 0
                '## Type 조회
                Pos1 = InStr(strRcvBuf, "rTYPE")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strBarno = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                    'strBarno = Val(strBarno)
                Else
                    '## NOTE: WorkNo가 전송되지 않은 에러처리
                    Exit Sub
                End If
                
                '## Level 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iQLEV")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strQCLevel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                
                '## QC 채널 조회
                Pos1 = 0: Pos2 = 0
                Pos1 = InStr(strRcvBuf, "iQFILE")
                If Pos1 > 0 Then
                    Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
                    strQCChannel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
                Else
                    '## NOTE: 바코드번호가 전송되지 않은 에러처리
                End If
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Rerun = ""
                    .Kind = "QC"
                    'If strOldBarno <> strBarno Then
                        'strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    'End If
                    
                    Call SetText(frmMain.spdOrder, strQCChannel, gRow, colPID)
                    Call SetText(frmMain.spdOrder, strQCLevel, gRow, colPNAME)
                End With
                          
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "m") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "m")
                    x2 = InStr(x1, strRcvBuf, GS)
        
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)

                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        'SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                'strQCData = GetQCResult_Detail(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    End If
                Loop
                
                
                
                x1 = 1
                Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                    x1 = InStr(x1, strRcvBuf, FS & "c")
                    x2 = InStr(x1, strRcvBuf, GS)
            
            '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                    'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                    strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                    x2 = x2 + 1
                    x1 = InStr(x2, strRcvBuf, GS)
                    strResult = Mid(strRcvBuf, x2, x1 - x2)
            
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        
                    End If
                Loop
                
                Exit Sub
            End If
            '#########  QC Define ##########################################################
        
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
        
        Next
        
        
    End With

End Sub

Private Sub Phase_Serial_PFA200()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case vbCr
                Call SerialRcvData_PFA200
                
                RcvBuffer = ""
                
                miLineNo = miLineNo + 1
                
            Case Is <> 10
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i

End Sub


Private Sub SerialRcvData_PFA200()
    Dim RS_L            As ADODB.Recordset
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
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        strRcvBuf = Replace(strRcvBuf, vbLf, "")
'Buffer = ""
'Buffer = Buffer & "PFA-100" & vbCrLf
'Buffer = Buffer & "REV. 2.20   S/N: 3954 " & vbCrLf
'Buffer = Buffer & "05/31/10       01:12 PM" & vbCrLf
'Buffer = Buffer & "ID#: 010000159846" & vbCrLf
'Buffer = Buffer & "Test Type: Collagen/ADP" & vbCrLf
'Buffer = Buffer & "SAMPLE  A:   114 SEC" & vbCrLf
'Buffer = Buffer & "cs: 6781" & vbCrLf

        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
        
        If UCase(Mid(strRcvBuf, 1, 3)) = "PFA" Then
            miLineNo = 1
            
        ElseIf Mid(strRcvBuf, 1, 3) = "ID#" Then
            strBarno = Trim(Mid(strRcvBuf, 5, 12))
            mResult.BarNo = strBarno
    
        ElseIf Mid(strRcvBuf, 1, 10) = "Test Type:" Then
            strIntBase = Trim(Mid(strRcvBuf, 11))
            mResult.IntBase = Trim(strIntBase)
            
        ElseIf Mid(strRcvBuf, 1, 3) = "QC:" Then
            mResult.Kind = "QC"
    
        ElseIf (Mid(strRcvBuf, 1, 9) = "SAMPLE A:") Or (Mid(strRcvBuf, 1, 9) = "SAMPLE B:") Then
            strResult = Mid(strRcvBuf, 10)
            If InStr(UCase(strRcvBuf), "SEC") = 0 Then
                strResult = Trim(strResult)
            Else
                varResult = Split(strResult, "Sec")
                strResult = Trim(varResult(0))
                strFlag = ""
                
                If Left(strResult, 1) = ">" And IsNumeric(Right(strResult, 1)) <> True Then
                    '비정상 결과 & Flag
                    strResult = Mid(strResult, 1, Len(strResult) - 1)
                End If
            End If
            
            mResult.RESULT = strResult
            
        Else
            If miLineNo = 7 And mResult.Kind <> "QC" Then
                If Trim(strRcvBuf) <> "" Then
                    strFlag = Trim(strRcvBuf)
                End If
            End If
            
            If UCase(Mid(strRcvBuf, 1, 3)) = "CS:" Or (miLineNo >= 7 And mResult.Kind <> "QC") Or (miLineNo >= 8 And mResult.Kind = "QC") Then
                strBarno = mResult.BarNo
                strIntBase = mResult.IntBase
                strResult = mResult.RESULT
                
                With mResult
                    .BarNo = strBarno
                    .SpcPos = strSeq
                    .Seq = strSeq
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                            
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
        
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
        
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            strState = "R"
                            
                            '-- BIORAD QC 저장
'                            If mResult.Kind = "QC" Then
'                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                Call SendBioRadQC(strQCData)
'                            End If
                    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH,QCTEMP " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strQCTemp = Trim(RS_L.Fields("QCTEMP"))
        
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
        
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            If strQCTemp = "1" Then
                                strResult = SetResult(strResult, strIntBase)
                            End If
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
        
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
        
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        
                    End If
                    
                End If
                            
                .spdResult.RowHeight(-1) = 14
                            
                '## DB에 결과저장
                If .optTrans(0).Value = True And strState = "R" Then
                    Res = SaveTransData_MCC(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "Failed", gRow, colSTATE
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
            End If
        End If
    End With

End Sub

Private Sub Phase_Serial_AFIAS6()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbCr
                Call SerialRcvData_AFIAS6
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub


Private Sub SerialRcvData_AFIAS6()
    Dim RS_L            As ADODB.Recordset
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
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strRcvBuf = strRecvData(intCnt)
            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
            strRackNo = ""
            strTubePos = ""
            strSeq = ""
                        
            With mResult
                .BarNo = strBarno
                .SpcPos = strSeq
                .Seq = strSeq
                .RackNo = strRackNo
                .TubePos = strTubePos
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            End With
            
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
            If gRow <= 0 Then
                Exit Sub
            End If
                        
            strIntBase = mGetP(strRcvBuf, 8, "|")
            strResult = mGetP(strRcvBuf, 11, "|")
                        
            If strIntBase <> "" And strResult <> "" Then
                If gPatOrdCd <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))

                        '-- 결과Row 추가
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If

                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE

                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- 결과 List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        strState = "R"
                        
                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                        
                    End If
                Else
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))

                        '-- 결과Row 추가
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If

                        '소수점 처리, 결과 형태 처리
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '진행상태 표시("결과")
                        SetText .spdOrder, "결과", gRow, colSTATE

                        '결과값 표시
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- 결과 List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                        
                        '-- 로컬 저장
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If

                        '-- 결과Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                    End If
                    
                End If
                
            End If
                        
            .spdResult.RowHeight(-1) = 14
                        
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
        Next
    End With

End Sub


Private Sub Phase_Serial_ADVIA1800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long


On Error GoTo RST

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        sRcvState = "": sSndState = ""
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case Else
                        intPhase = 1
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case STX
'                        intBufCnt = 1
'                        Erase strRecvData
'                        ReDim Preserve strRecvData(intBufCnt)
                        RcvBuffer = ""
                    Case EOT
                        Select Case sRcvState
                            Case "Q"
                                intPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                frmMain.comEqp.Output = ENQ
                                sSndState = "E"
                                
                            Case "R"
                                intPhase = 1
                        End Select
                        
                        sRcvState = ""
                    
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbLf
                        intPhase = 2
                        If RcvBuffer <> "" Then
                            Call SerialRcvData_ADVIA1800
                            RcvBuffer = ""
                        End If
                        
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbCr
                    Case ETB
                    
                    Case Else
                        intPhase = 2
'                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        RcvBuffer = RcvBuffer & BufChar
                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case ACK
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                                
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket
                                intPhase = 1
                        End Select
                    
                    Case ENQ
                        'Erase strRecvData
                        RcvBuffer = ""
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case NAK
                        Select Case sSndState
                            Case "E"
                                frmMain.comEqp.Output = Chr(5)
                                intPhase = 3
                            Case "P"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                            Case "L"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        'Erase strRecvData
                        RcvBuffer = ""
                        intPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                End Select
        End Select
    Next i
                
    Exit Sub
    
RST:
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    strErrMsg = strErrMsg & "ORDER    : " & mOrder.BarNo & vbNewLine
    strErrMsg = strErrMsg & "RESLLT   : " & mResult.BarNo & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
            
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
Private Sub GetOrder_ADVIA1800(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
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
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ADVIA1800(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub


'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ADVIA1800(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ADVIA1800 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' " 89M 81M 82M 90M 91M108M 85M"
            If Trim(AdoRs_Local.Fields("SENDCHANNEL").Value) <> "" Then
                strExamCode = strExamCode & Right(Space(3) & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 3) & "M"
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA1800 = strExamCode
    
End Function



Private Sub SerialRcvData_ADVIA1800()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strTmp          As String
    Dim i               As Integer
    Dim iBCpos          As Integer
    
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim strKind     As String
    Dim iPos        As Integer
    
    Dim varIntBase()    As String
    Dim varResult()     As String
    Dim varFlag()       As String
    
    Dim strUseRes       As String
    Dim strCREA         As String
    Dim strTP           As String

On Error GoTo RST
    
    
    iBCpos = 2
    
    With frmMain
        'For intCnt = 1 To UBound(strRecvData)
        '    strRcvBuf = strRecvData(intCnt)
            
            strRcvBuf = RcvBuffer
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, iBCpos, 1)
            
            Select Case strType
                Case "q"    '## Request Information(Batch)
                    sRcvState = "Q"
                    sSndState = ""
                    
                Case "Q"    '## Request Information
                    sRcvState = "Q"
                    sSndState = ""
                
                    iTmpPendingFlag = Val(Mid$(strRcvBuf, iBCpos + 6, 2))
                    iPendingFlag = iPendingFlag + iTmpPendingFlag
                    
                    For i = 1 To iPendingFlag
                        strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 9 + 13 * (i - 1), 13))
                        
                        With mOrder
                            .NoOrder = False
                            .BarNo = strBarno
                        End With
                        
                        Call GetOrder_ADVIA1800(strBarno, gHOSP.RSTTYPE)
                        Call SendOrder_ADVIA1800
                    Next
                
                Case "R"
                    sRcvState = "R"
                    
                    iTBlockNo = Val(Mid$(strRcvBuf, iBCpos + 2, 2))
                    iCBlockNo = Val(Mid$(strRcvBuf, iBCpos + 4, 2))
                    iItemNo = Val(Mid$(strRcvBuf, iBCpos + 6, 3))
                    
                    iBCpos = iBCpos + 6
                    
                    strKind = Mid$(strRcvBuf, iBCpos + 17, 1)       'N:Sample, C:Control
                    strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 19, 13))
                                    
                    strTemp2 = Trim$(Mid$(strRcvBuf, iBCpos + 32, 7))
                    iPos = InStr(strTemp2, "-")
                             
                    If iPos = 0 Then
                        strRackNo = ""
                        strTubePos = ""
                    Else
                        strRackNo = Mid$(strTemp2, 1, iPos - 1)
                        strTubePos = Mid$(strTemp2, iPos + 1)
                    End If
                    
                    If strKind = "C" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    ReDim Preserve varIntBase(iItemNo)
                    ReDim Preserve varResult(iItemNo)
                    ReDim Preserve varFlag(iItemNo)
                    
                    If iCBlockNo = 1 Then
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 89 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    Else
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 39 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    End If
                    
                    If mResult.Rerun = "R" Then       'Rerun Result
                        mResult.Kind = mResult.Kind & "R"
                    End If
                    
                    strCREA = ""
                    strTP = ""
                    
                    For i = 1 To iItemNo
                        strIntBase = varIntBase(i)
                        strResult = varResult(i)
                        
'91  C3750N3 CREA(단회)
'110 C2200-1 TP(단회뇨)
                        If strIntBase = "91" Then
                            strCREA = strResult
                        End If
                        If strIntBase = "110" Then
                            strTP = strResult
                        End If

RST1:
                        If strIntBase <> "" And strResult <> "" And strResult <> "ERROR" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                    lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte")) & ""
                                    
                                    
                                    'LDH  체액
                                    'If lsTestCode = "C2590N1" Or lsTestCode = "C2590N2" Then '혈액 Or lsTestCode = "B2590"
                                    If lsTestCode = "B2590N1" Or lsTestCode = "B2590N2" Then '혈액 Or lsTestCode = "B2590"
                                        If IsNumeric(strResult) Then
                                            strResult = strResult / 6
                                        End If
                                    End If
                                    
                                    '-- 결과Row 추가
                                    lsRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < lsRstRow Then
                                        .spdResult.MaxRows = lsRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    'CRR 적용
                                    If strKind <> "QC" Then
                                        strResult = getCRRValue(lsTestCode, strResult)
                                    End If
                                    
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                    
                                    '-- High, Low 색 표시
                                    If strJudge <> "" Then
                                        SetForeColor .spdResult, lsRstRow, lsRstRow, colRMACHRESULT, colRLISRESULT, 255, 0, 0
                                    End If
                                                                        
                                    '-- 로컬 저장
                                    SetLocalDB gRow, lsRstRow, "1", ""
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                    End If
                                    
                                    strState = "R"
                                    
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                    
                                    If lsTestCode = "C3730N1" Or lsTestCode = "C3750" Or lsTestCode = "C7230" Or lsTestCode = "C3750N3" Or lsTestCode = "C2302N6" Then
                                        Call CalProcess(spdOrder, spdResult, lsTestCode)
                                    End If

                                End If
                            Else
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                    strQCTemp = Trim(RS_L.Fields("QCTEMP")) & ""
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                    
                                    'LDH  체액
                                    'If lsTestCode = "C2590N1" Or lsTestCode = "C2590N2" Then '혈액LDH Or lsTestCode = "B2590"
                                    If lsTestCode = "B2590N1" Or lsTestCode = "B2590N2" Then '혈액LDH Or lsTestCode = "B2590"
                                        If IsNumeric(strResult) Then
                                            strResult = strResult / 6
                                        End If
                                    End If
                                    
                                    '-- 결과Row 추가
                                    lsRstRow = .spdResult.DataRowCnt + 1
                                    If .spdResult.MaxRows < lsRstRow Then
                                        .spdResult.MaxRows = lsRstRow
                                    End If
            
                                    '소수점 처리, 결과 형태 처리
                                    strMachResult = strResult
                                    If strQCTemp = "1" Then
                                        strResult = SetResult(strResult, strIntBase)
                                    End If
                                    strJudge = SetJudge(strResult, strIntBase)
                                    
                                    'CRR 적용
                                    If strKind <> "QC" Then
                                        strResult = getCRRValue(lsTestCode, strResult)
                                    End If
                                    
                                    '진행상태 표시("결과")
                                    SetText .spdOrder, "결과", gRow, colSTATE
            
                                    '결과값 표시
                                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                            SetText .spdOrder, strResult, gRow, intCol
                                            Exit For
                                        End If
                                    Next
            
                                    '-- 결과 List
                                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                                    
                                    '-- High, Low 색 표시
                                    If strJudge <> "" Then
                                        SetForeColor .spdResult, lsRstRow, lsRstRow, colRMACHRESULT, colRLISRESULT, 255, 0, 0
                                    End If
                                    
                                    '-- 로컬 저장
                                    SetLocalDB gRow, lsRstRow, "1", ""
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        
                                        strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                        
                                    End If
                                    
                                    If strState <> "R" Then
                                        strState = ""
                                    End If
            
                                    '-- 결과Count
                                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                        SetText .spdOrder, "1", gRow, colRCNT
                                    Else
                                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                    End If
                                End If
                                
                            End If
                            
                        End If
                    Next
                    
                    If strTP <> "" And strCREA <> "" And IsNumeric(strTP) And IsNumeric(strCREA) Then
                        strIntBase = "PRCR"
                        strResult = strTP / strCREA
                        strResult = Format(strResult, "#,##0.00")
                        strTP = ""
                        strCREA = ""
                        GoTo RST1
                    End If
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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
                            
                            If lsTestCode = "C3730N1" Or lsTestCode = "C3750" Or lsTestCode = "C7230" Or lsTestCode = "C3750N3" Or lsTestCode = "C2302N6" Then
                                Call CalProcess(spdOrder, spdResult, lsTestCode)
                            End If
                            
                        End If
                        strState = ""
                    End If
            End Select
        'Next
    End With

    Exit Sub
    
RST:
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_SerialRcvData_ADVIA1800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    strErrMsg = strErrMsg & "ORDER    : " & mOrder.BarNo & vbNewLine
    strErrMsg = strErrMsg & "RESLLT   : " & mResult.BarNo & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
End Sub


Private Sub SerialRcvData_RAPIDPOINT500()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String

    Dim x   As Integer
    Dim c   As Integer
    Dim MsgID   As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String
    Dim iQID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
'    Dim strQCResult As String
    
    Dim iQLEV$, iQLOT$, strAnalyte$
    Dim db_tmp As String * 100
    
    Dim Pos1            As Long
    Dim Pos2            As Long
    Dim strQCChannel    As String
    
    With frmMain
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", RcvBuffer, "A")
        End If
        '-- 테스트용 -----------------
        
        x = InStr(1, RcvBuffer, FS)
        If RcvBuffer <> "" Then
            MsgID = Mid(RcvBuffer, 2, x - 2)
        End If
        Select Case MsgID
            Case "ID_REQ"
                Call SendMessage_1200("ID_DATA")
            Case "SMP_START"
            Case "SMP_NEW_AV"
                Do Until x = 0
                    x = InStr(x, RcvBuffer, "r")
                    If x = 0 Then Exit Do
                    If Mid(RcvBuffer, x, 4) = "rSEQ" Then
                        x = x + 5
                        c = InStr(x, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, x, c - x)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("SMP_REQ")
                Loop
            
            Case "SYS_READY"
            Case "SYS_NOT_READY"
            Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
                GoTo RST
            Case "CAL_ABORT"
            Case "QC_START"
            Case "QC_NEW_AV"
                Do Until x = 0
                    x = InStr(x, RcvBuffer, "r")
                    If x = 0 Then Exit Do
                    If Mid(RcvBuffer, x, 4) = "rSEQ" Then
                        x = x + 5
                        c = InStr(x, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, x, c - x)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("QC_REQ")
                Loop
            Case "QC_NEW_DATA", "QC_EDIT_DATA"
                GoTo RST
        End Select
            
        Exit Sub

RST:
        MsgBuf = RcvBuffer
        
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
            strSeq = Trim(rSeq)
            
            If strBarno = "" Or Not IsNumeric(strBarno) Then
                Exit Sub
            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                If strOldBarno <> strBarno Then
                    strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                End If
            End With
            
            strState = "O"
                        
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
                Debug.Print strIntBase
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
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
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
                
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Loop
            
            .spdResult.RowHeight(-1) = 14
            
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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

            
        '>> If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
        
        ElseIf MsgID = "QC_NEW_DATA" Or MsgID = "QC_EDIT_DATA" Then
            '## Type 조회
            Pos1 = InStr(MsgBuf, "rTYPE")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strBarno = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
                'strBarno = Val(strBarno)
            Else
                '## NOTE: WorkNo가 전송되지 않은 에러처리
                Exit Sub
            End If

            '## Level 조회
            Pos1 = 0: Pos2 = 0
            Pos1 = InStr(MsgBuf, "iQLEV")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strQCLevel = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
            Else
                '## NOTE: 바코드번호가 전송되지 않은 에러처리
            End If


            '## QC 채널 조회
            Pos1 = 0: Pos2 = 0
            Pos1 = InStr(MsgBuf, "iQFILE")
            If Pos1 > 0 Then
                Pos2 = InStr(Mid$(MsgBuf, Pos1), FS)
                strQCChannel = mGetP(Mid$(MsgBuf, Pos1, Pos2), 2, GS)
            Else
                '## NOTE: 바코드번호가 전송되지 않은 에러처리
            End If
            
            strQCChannel = strQCLevel
            
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "aMod") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                aMod = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'iIID
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "iIID") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                iIID = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'rSEQ
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                rSeq = Mid(MsgBuf, x1, x2 - x1)
'            End If
'
'            'PID
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "iPID") + 5
'            If x1 <> 5 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                iPID = Mid(MsgBuf, x1, x2 - x1)
'            End If
'            'DATE
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rDATE") + 6
'            If x1 <> 6 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                sRstDate = Mid(MsgBuf, x1, x2 - x1)
'                sRstDate = ConvertDateType(sRstDate)
'            End If
'            'TIME
'            x1 = 1
'            x1 = InStr(x1, MsgBuf, "rTIME") + 6
'            If x1 <> 6 Then
'                x2 = InStr(x1, MsgBuf, GS)
'                sRstTime = Mid(MsgBuf, x1, x2 - x1)
'                sRstTime = Format(sRstTime, "HHNNSS")
'            End If
'
'            x2 = 0
'
'            '접수번호, SeqNo
'            strBarno = Trim(iPID)
'            strSeq = Trim(rSeq)
'
'            If strBarno = "" Or Not IsNumeric(strBarno) Then
'                Exit Sub
'            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                .Kind = "QC"
                'If strOldBarno <> strBarno Then
                    'strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                'End If
                    
                Call SetText(frmMain.spdOrder, strQCChannel, gRow, colPID)
                Call SetText(frmMain.spdOrder, strQCLevel, gRow, colPNAME)
                
            End With
            
            strState = "O"
                        
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
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                'strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                strQCData = GetQCResult_Detail_Type2(gHOSP.LABCD, strQCChannel, strQCAnalyte, strResult)
                                
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
            Loop
                    
                    
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                x1 = InStr(x1, strRcvBuf, FS & "c")
                x2 = InStr(x1, strRcvBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, strRcvBuf, GS)
                strResult = Mid(strRcvBuf, x2, x1 - x2)
                
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                            lsTestName = Trim(RS_L.Fields("TESTNAME"))
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                            
                            strState = "R"
                            
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    Else
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                            lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                            lsSeqNo = Trim(RS_L.Fields("SEQNO"))
    
                            strQCLab = Trim(RS_L.Fields("QCLab") & "")
                            strQCLot = Trim(RS_L.Fields("QCLot") & "")
                            strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                            strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                            strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                            strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                            strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                            strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
    
                            '-- 결과Row 추가
                            lsRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < lsRstRow Then
                                .spdResult.MaxRows = lsRstRow
                            End If
    
                            '소수점 처리, 결과 형태 처리
                            strMachResult = strResult
                            strResult = SetResult(strResult, strIntBase)
                            strJudge = SetJudge(strResult, strIntBase)
                            
                            '진행상태 표시("결과")
                            SetText .spdOrder, "결과", gRow, colSTATE
    
                            '결과값 표시
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                            SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                            SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                            SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                            SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                            SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                            SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                            SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                            
                            '-- 로컬 저장
                            SetLocalDB gRow, lsRstRow, "1", ""
                            
                            '-- BIORAD QC 저장
                            If mResult.Kind = "QC" Then
                                
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                                
                            End If
                            
                            If strState <> "R" Then
                                strState = ""
                            End If
    
                            '-- 결과Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                    End If
                End If
                
            Loop
            
            '## DB에 결과저장
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- 저장 실패
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
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
            
            .spdResult.RowHeight(-1) = 14
            
            
        End If
        
        
        
    End With

End Sub

Private Sub SendMessage_1200(ByVal MsgHead As String)
    Dim chksum As Integer
    Dim Buffer As String
    Dim c As Integer
'    Dim R As Integer

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
            
        Case "QC_REQ"
            Buffer = STX & "QC_REQ" & FS & R_S & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & R_S & ETX
            
        Case "SMP_ORD"
    End Select
        
    For c = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, c, 1))
    Next c
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    frmMain.comEqp.Output = sSendData
    
    SetRawData "[Tx]" & sSendData
    
End Sub

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


Private Sub Phase_Serial_RAPIDPOINT500()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                AckOn = False
                RcvBuffer = BufChar
                
            Case EOT
                If AckOn = False Then
                    frmMain.comEqp.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    Call SerialRcvData_RAPIDPOINT500
                End If
            
            Case ACK
                AckOn = True
                RcvBuffer = RcvBuffer & BufChar
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                
        End Select
    Next i
            
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_ACLTOP()


    Dim strOutput   As String     '송신할 데이터
    Dim blnLast     As Boolean
    Dim intRow      As Integer
    Dim strBarno    As String
    Dim strItems    As String

    blnLast = False

    With frmMain.spdOrder
        If intSndPhase <= 3 Then
            For intRow = 1 To .DataRowCnt
                If GetText(frmMain.spdOrder, intRow, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, intRow, colSTATE) = "오더준비" Then
                    strBarno = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))
                    strItems = Trim(GetText(frmMain.spdOrder, intRow, colKEY1))
                    If intSndPhase = 3 Then
                        .Row = intRow
                        .Col = colCHECKBOX: .Text = "0"
                        .Col = colSTATE:    .Text = "오더전송"

                        If intRow = .DataRowCnt Then
                            blnLast = True
                        End If

                    End If
                    Exit For
                End If
            Next
        End If
    End With

    If intRow = frmMain.spdOrder.DataRowCnt Then
        blnLast = True
    End If

    Select Case intSndPhase
        Case 1  '## Header
        '''''            strOutput = "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr
            strOutput = intFrameNo & "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr & ETB
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
'''''        strOutput = strOutput & "P|" & mPNo & "||||^||||||||" & vbCr
            strOutput = intFrameNo & "P|" & mPNo & "||||^||||||||" & vbCr & ETB
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
            mPNo = mPNo + 1

        Case 3  '## Order
            '## 최초 보낼때
            If mOrder.IsSending = False Then
'''''         = strOutput & "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q" & vbCr
                strOutput = "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q"
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 3
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETB
                    If blnLast = True Then
                        intSndPhase = 4
                    Else
                        intSndPhase = 2
                    End If
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
                    strOutput = intFrameNo & strOutput & vbCr & ETB
                    If blnLast = True Then
                        intSndPhase = 4
                    Else
                        intSndPhase = 2
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
'''''            strOutput = strOutput & "L|1|N"
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select


    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub

Private Sub Phase_Serial_ACLTOP()
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
                        intBufCnt = 0
                        Erase strRecvData
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_ACLTOP
                        Else
                            frmMain.comEqp.Output = ACK
                            SetRawData "[Tx]" & ACK
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                        Else
                            intBufCnt = intBufCnt + 1
                        End If
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    Case ETX
                        intPhase = 3
                    Case EOT
                        intPhase = 1
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        Call SerialRcvData_ACLTOP
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i

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
Private Sub GetOrder_ACLTOP(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String

    intRow = -1

    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
                    intRow = i
                    Exit For
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If

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

        .spdOrder.RowHeight(-1) = 12

        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_ACLTOP(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더준비", intRow, colSTATE)
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
        End If

        SetText frmMain.spdOrder, "1", intRow, colCHECKBOX

        '-- 현재 Row
        gRow = intRow

    End With

End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_ACLTOP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ACLTOP = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "@^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_ACLTOP = Mid(strExamCode, 2)

End Function



Private Sub SerialRcvData_ACLTOP()
    Dim RS_L            As ADODB.Recordset
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
    Dim strAspect       As String

    Dim strTemp1        As String
    Dim strTemp2        As String

    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq

    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    Dim varBarno        As Variant
    Dim i               As Integer

    Dim strUseRes       As String

    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)

            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"    '## Header
                    '1H|@^\|<1504128210_21570><1504128210_21571>||acl|||||LIS||P|1394-97|20170830172330
                    mOrder.MsgID = Trim(mGetP(strRcvBuf, 3, "|"))
                    mOrder.Sender = Trim(mGetP(strRcvBuf, 5, "|"))
                    mOrder.Receiver = Trim(mGetP(strRcvBuf, 10, "|"))
                    mOrder.Version = Trim(mGetP(strRcvBuf, 13, "|"))

                Case "P"    '## Patient
                Case "Q"    '## Request Information

                    'Q|1|^1001@^1002@^1003@^1004@^1005@^1006@^1008||||||||||O@N
                    'Q|1|^198772||||||||||O@N
                    'Q|1|^1310250941@^1310250867||||||||||O@N


                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strTemp1 = Replace(strTemp1, "^", "")

                    varBarno = Split(strTemp1, "@")

                    For i = 0 To UBound(varBarno)
                        mOrder.BarNo = varBarno(i)
                        Call GetOrder_ACLTOP(varBarno(i), gHOSP.RSTTYPE)
                    Next

'                    With mOrder
'                        .NoOrder = False
'                        .BarNo = strBarno
'                        .Seq = mGetP(strTemp1, 3, "^")
'                        .RackNo = mGetP(strTemp1, 4, "^")
'                        .TubePos = mGetP(strTemp1, 5, "^")
'                    End With

                   ' Call GetOrder(strBarno, gHOSP.RSTTYPE)
                    strState = "Q"
                    mPNo = 1

                Case "O"
                    strBarno = mGetP(strRcvBuf, 3, "|")

                    With mResult
                        .BarNo = strBarno
                        '.SpcPos = strTubePos & "/" & strRackNo
                        '.Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))

                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)

                        End If
                    End With

                Case "R"
                    'R|1|^^^131|28.4|s||N||F@V||SysAdmin^SysAdmin||20170826150315|
                    'R|1|^^^541|103.6|D mAbs||N||F@V||SysAdmin^SysAdmin||2017090108
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    If strIntBase = "131" Then
                        strIntBase = strIntBase & UCase(mGetP(strRcvBuf, 5, "|"))
                    End If

                    'R|1|^^^2241|0.3|microg/mLFEU||N||F@V||SysAdmin^SysAdmin||2017
                    'R|2|^^^2241|172|ng/mL||N||F@V||SysAdmin^SysAdmin||20170901083115|acl^03^2

                    ' D-Dimer
                    If strIntBase = "2241" Then
                        If mGetP(strRcvBuf, 5, "|") = "microg/mLFEU" Then
                            strIntBase = strIntBase
                        Else
                            strIntBase = ""
                        End If
                    End If

                    strResult = mGetP(strRcvBuf, 4, "|")

                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")

                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)

                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""

                                '-- BIORAD QC 저장
                                If Mid(strBarno, 1, 2) = "QC" Then
                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                End If


                                strState = "R"

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If

                            End If
                        Else
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "

                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""

                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")

                                '-- 결과Row 추가
                                lsRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < lsRstRow Then
                                    .spdResult.MaxRows = lsRstRow
                                End If

                                '-- 소수점 처리, 결과 형태 처리
                                If strUseRes <> "" Then
                                    strMachResult = strResult
                                    strResult = SetResult(strResult, strIntBase)
                                End If
                                strJudge = SetJudge(strResult, strIntBase)


                                '진행상태 표시("결과")
                                SetText .spdOrder, "결과", gRow, colSTATE

                                '결과값 표시
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next

                                '-- 결과 List
                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치

                                '-- 로컬 저장
                                SetLocalDB gRow, lsRstRow, "1", ""

                                '-- BIORAD QC 저장
                                If Mid(strBarno, 1, 2) = "QC" Then
                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                                End If

                                If strState <> "R" Then
                                    strState = ""
                                End If

                                '-- 결과Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                            End If

                        End If

                    End If

                    .spdResult.RowHeight(-1) = 14

                Case "C"    '## Comment
                    '## Abnormal 결과일때 Comment 저장
                    If strFlag <> "N" Then
                        strTemp1 = mGetP(strRcvBuf, 4, "|")
                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
                    End If

                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_MCC(gRow)

                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
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

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            pBuffer = comEqp.Input
            
            dtpToday = Now
            
            If fraInterface.Visible = False Then
                tmrComm.Interval = 20000
                tmrComm.Enabled = True
                
                tmrFlipFlop.Interval = 500
                tmrFlipFlop.Enabled = True
                
                lblCommStatus.Caption = "장비 검사결과가 수신되었습니다. 인터페이스 창에서 확인하세요!"
            End If
            
            'txtRcv.Text = pBuffer
            SetRawData "[Rx]" & pBuffer

            
            Select Case UCase(gHOSP.MACHNM)
                ' 타이머를 사용해야 해서 메인에서 처리함
                Case "ADVIA2120-1", "ADVIA2120-2"
                        Call Phase_Serial_ADVIA2120
                        
                ' 타이머를 사용해야 해서 메인에서 처리함
                Case "CT500"
                        Call Phase_Serial_CT500
                        
                Case "VERSACELL"
                        Call Phase_Serial_VERSACELL
                
                Case "RAPIDLAB348"
                        Call Phase_Serial_RAPIDLAB348
                
                Case "PFA200"
                        Call Phase_Serial_PFA200
                
                Case "AFIAS6"
                        Call Phase_Serial_AFIAS6
                
                Case "ADVIA1800-1", "ADVIA1800-2"
                        Call Phase_Serial_ADVIA1800
                
                Case "RAPIDPOINT500"
                        Call Phase_Serial_RAPIDPOINT500
                
                Case "ACLTOP"
                        Call Phase_Serial_ACLTOP
                
                Case "VESCUBE"
                        Call Phase_Serial_VESCUBE
                
                Case "OSMOPRO"
                        Call Phase_Serial_OSMOPRO
                
                Case "URINSCANPRO"
                    lngBufLen = Len(pBuffer)
                
                    For i = 1 To lngBufLen
                        BufChar = Mid$(pBuffer, i, 1)
                        Select Case intPhase
                            Case 1
                                Select Case BufChar
                                    Case STX
                                        RcvBuffer = ""
                                        intPhase = 2
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                            Case 2
                            
                                Select Case BufChar
                                    Case ETX
                                        Call SerialRcvData_UrinscanPro
                                        RcvBuffer = ""
                                        intPhase = 1
                                    Case Else
                                        RcvBuffer = RcvBuffer & BufChar
                                End Select
                        End Select
                    Next i
                
                Case "ACLELITE"
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
                                        comEqp.Output = ACK
                                        SetRawData "[Tx]" & ACK
                                    Case ACK
                                        '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                        '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                        strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                        dtpToday.Value = Format(strDate, "####-##-##")
                                        
                                        DoEvents
                                        
                                        If strState = "Q" Then Call SendOrder_ACLELITE
                                
                                End Select
                            Case 2      '## Transfer Phase
                                Select Case BufChar
                                    Case ENQ
                                        Erase strRecvData
                                        comEqp.Output = ACK
                                        SetRawData "[Tx]" & ACK
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
                                        comEqp.Output = ACK
                                        SetRawData "[Tx]" & ACK
                                End Select
                            Case 4      '## Termination Phase
                                Select Case BufChar
                                    Case STX
                                        intPhase = 2
                                    Case EOT
                                        '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                        '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                        strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                        dtpToday.Value = Format(strDate, "####-##-##")
        
                                        DoEvents
                                        
                                        Call SerialRcvData_ACLELITE
                                        
                                        If strState = "Q" Then
                                            intSndPhase = 1
                                            intFrameNo = 1
                                            comEqp.Output = ENQ
                                            SetRawData "[Tx]" & ENQ
                                        End If
                                        
                                        intPhase = 1
                                End Select
                        End Select
                    Next i
                                    
                Case Else
                        Call Serial_Protocol
                        
            End Select
                        
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
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

End Sub


Private Sub Form_Load()

On Error GoTo RST

    Me.Width = 20940
    Me.Height = 12585
    
    lblHospInfo.Caption = gHOSP.HOSPNM & "  " & gHOSP.MACHNM & "  " & gHOSP.USERNM & "[" & gHOSP.USERID & "]" '& "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Me.Caption = gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
    
    Call CtlInitializing
        
    '-- Menu Set
    Call SetMenu
    
    '-- 검사코드
    Call GetTestList
    
    '-- 오더코드
    Call GetOrderMST

    '-- 검사명 보이기
    Call SetExamCode
    

    '-- 통신설정
    Call GetCommList

    If gComm.COMTYPE = "1" Then
        comEqp.CommPort = gComm.COMPORT
        comEqp.RTSEnable = gComm.RTSEnable
        comEqp.DTREnable = gComm.DTREnable
        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

        If comEqp.PortOpen = False Then
            comEqp.PortOpen = True
        End If

        If comEqp.PortOpen Then
            lblStatus.Caption = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            lblStatus.Caption = "통신포트에 연결 되지 않았습니다"
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    Else
        If gComm.TCPTYPE = "1" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen

            lblStatus.Caption = "TCP " & gComm.TCPPORT & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

            lblStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " 포트에 연결 되었습니다"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        End If
    End If

    If gHOSP.MACHNM = "ADVIA2120-1" Or gHOSP.MACHNM = "ADVIA2120-2" Then
        cmdInit.Visible = True
        Call InitialComm
    Else
        cmdInit.Visible = False
    End If
        
    frame1.Visible = True
    frame1.ZOrder 0

    
    '변수 초기화(Advia1650)
    iPendingFlag = 0: iTotQueryFlag = 0: iTmpPendingFlag = 0: iIdleFlag = 0
    iOrderFlag = 0: iResultFlag = 0
    sRcvState = "": sSndState = ""
    
'    txtURL.Text = gHOSP.ALLEREIPATH
'
'    tmrResult.Interval = 1000
'    tmrResult.Enabled = True

'spdOrder.MaxRows = 10

    
    Exit Sub
    
RST:
    frame1.Visible = True
    frame1.ZOrder 0
    
    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    Else
        MsgBox Err.Number & vbNewLine & Err.Description
    End If
    
End Sub

'-- 검사마스터 조회
Public Sub GetCommList()
    Dim i As Integer
    Dim Ret As Integer
    
    If gComm.COMTYPE = "1" Then
        optComType(0).Value = True
        frameCom.Enabled = True
        frameTCP.Enabled = False
    Else
        optComType(1).Value = True
        frameCom.Enabled = False
        frameTCP.Enabled = True
    End If
    
    Ret = -1
    For i = 0 To cboPort.ListCount - 1
        If gComm.COMPORT = Trim(cboPort.List(i)) Then
            cboPort.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    
'    If Ret = -1 Then
'        cboPort.ListIndex = 1
'    End If
    
    Ret = -1
    For i = 0 To cboBaudrate.ListCount - 1
        If gComm.SPEED = Trim(cboBaudrate.List(i)) Then
            cboBaudrate.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboBaudrate.ListIndex = 4
    End If
    
    Ret = -1
    For i = 0 To cboDatabit.ListCount - 1
        If gComm.DATABIT = Trim(cboDatabit.List(i)) Then
            cboDatabit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboBaudrate.ListIndex = 1
    End If

    Ret = -1
    For i = 0 To cboStartbit.ListCount - 1
        If gComm.STARTBIT = Trim(cboStartbit.List(i)) Then
            cboStartbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboStartbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To cboStopbit.ListCount - 1
        If gComm.STOPBIT = Trim(cboStopbit.List(i)) Then
            cboStopbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboStopbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To cboParity.ListCount - 1
        If gComm.Parity = Trim(cboParity.List(i)) Then
            cboParity.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        cboParity.ListIndex = 0
    End If
    
    '--------------------------------------------
    
    If gComm.TCPTYPE = "1" Then
        optTCPType(0).Value = True
    Else
        optTCPType(1).Value = True
    End If
    
    txtTCPIP.Text = gComm.TCPIP
    txtTCPPort.Text = gComm.TCPPORT
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub
    
    '-- 인터페이스
    frame1.Width = Me.ScaleWidth - 150
    frame1.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdOrder.Width = Me.ScaleWidth - spdResult.Width - 400
    spdOrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    spdResult.Left = spdOrder.Left + spdOrder.Width + 50
    spdResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 결과조회
    frame2.Width = Me.ScaleWidth - 150
    frame2.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdROrder.Width = Me.ScaleWidth - spdRResult.Width - 500
    spdROrder.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    spdRResult.Left = spdOrder.Left + spdROrder.Width + 50
    spdRResult.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 검사설정
    frame3.Width = Me.ScaleWidth - 150
    frame3.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150
    
    spdTest.Width = Me.ScaleWidth - frameTestSet.Width - 600
    spdTest.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500
    
    frameTestSet.Left = spdTest.Left + spdTest.Width + 50
    frameTestSet.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 500

    '-- 통신설정
    Frame4.Width = Me.ScaleWidth - 150
    Frame4.Height = Me.ScaleHeight - (Picture1.Height + Picture2.Height) - 150

End Sub





Private Sub fraInterface_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
End Sub

Private Sub frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblComSave.ForeColor = vbBlack
    lblTcpSave.ForeColor = vbBlack
    
    shpCom.BorderColor = &H808080
    shpTcp.BorderColor = &H808080
    
    
End Sub

Private Sub frameTestSet_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblActionTest(i).ForeColor = vbBlack
        shpA(i).BorderColor = &H808080
    Next
    
End Sub

Private Sub imgDelete_Click()
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Trim(txtEqpCD.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtOChannel.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    Set Test_Property = New Scripting.Dictionary

    With Test_Property
        .Add "EQPCD", txtEqpCD.Text
        .Add "OCH", txtOChannel.Text
        .Add "TESTCD", txtTestCd.Text
    End With
    
    Set objTest_Property = New clsCommon
    
    With objTest_Property
        .SetAdoCn AdoCn_Local
        If .DelTestInfo(Test_Property) Then
            '-- 저장 오류
            Call GetTestList
        Else
            '-- 저장 오류
            Call GetTestList
        End If
    End With

End Sub

Private Sub imgSave_Click()
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If Trim(txtEqpCD.Text) = "" Then
        MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtOChannel.Text) = "" Then
        MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
        txtOChannel.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRChannel.Text) = "" Then
        MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
        txtRChannel.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTestCd.Text) = "" Then
        MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
        txtTestCd.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTestNm.Text) = "" Then
        MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
        txtTestNm.SetFocus
        Exit Sub
    End If
    
    
    Set Test_Property = New Scripting.Dictionary

    With Test_Property
        .Add "EQPCD", txtEqpCD.Text
        .Add "SEQ", txtSeq.Text
        .Add "OCH", txtOChannel.Text
        .Add "RCH", txtRChannel.Text
        .Add "TESTCD", txtTestCd.Text
        .Add "TESTNM", txtTestNm.Text
        .Add "ABBRNM", txtAbbrNm.Text
        .Add "RES", txtResSpec.Text
        .Add "REFL", txtRefLow.Text
        .Add "REFH", txtRefHigh.Text
        .Add "RSTTYPE", cboResultType.Text
        If optCutUse(0).Value = True Then
            .Add "CUTUSE", "Y"
        Else
            .Add "CUTUSE", "N"
        End If
        .Add "COLIN", txtCOLIn.Text
        .Add "COLCP", cboCOL.Text
        .Add "COLOUT", txtCOLOut.Text
        .Add "COHIN", txtCOHIn.Text
        .Add "COHCP", cboCOH.Text
        .Add "COHOUT", txtCOHOut.Text
        .Add "COMOUT", txtCOMOut.Text
    End With
    
    Set objTest_Property = New clsCommon
    
    With objTest_Property
        .SetAdoCn AdoCn_Local
        If .LetTestInfo(Test_Property) Then
            '-- 저장 오류
            Call GetTestList
        Else
            '-- 저장 오류
            Call GetTestList
        End If
    End With

End Sub



Public Sub CtlInitializing()
    Dim intComPortExist As Long
    Dim i As Integer
    
    frame1.Left = 50
    frame1.Top = 1650
    
    frame2.Left = 50
    frame2.Top = 1650
    
    frame3.Left = 50
    frame3.Top = 1650
    
    Frame4.Left = 50
    Frame4.Top = 1650
    
    dtpToday.Value = Now
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    
    '-- 인터페이스
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    
    '-- 장비결과
    spdROrder.MaxRows = 0
    spdRResult.MaxRows = 0
        
    '-- 검사코드 설정
    spdTest.MaxRows = 0
    
    cboCOL.AddItem "<"
    cboCOL.AddItem "<="
    cboCOL.ListIndex = 0
    
    cboCOH.AddItem ">"
    cboCOH.AddItem ">="
    cboCOH.ListIndex = 0
    
    cboResultType.AddItem "변함없음"
    cboResultType.AddItem "정량"
    cboResultType.AddItem "정성"
    cboResultType.AddItem "정량(정성)"
    cboResultType.AddItem "정성(정량)"
    cboResultType.ListIndex = 0
    
    txtEqpCD.Text = gHOSP.HOSPCD
    
    '-- 통신설정
    cboPort.AddItem ("1")
    cboPort.AddItem ("2")
    cboPort.AddItem ("3")
    cboPort.AddItem ("4")
    cboPort.AddItem ("5")
    cboPort.AddItem ("6")
    cboPort.AddItem ("7")
    cboPort.AddItem ("8")
    cboPort.AddItem ("9")
    cboPort.AddItem ("10")
    cboPort.AddItem ("11")
    cboPort.AddItem ("12")
    cboPort.AddItem ("13")
    cboPort.AddItem ("14")
    cboPort.AddItem ("15")
    cboPort.AddItem ("16")
    
    cboPort.Clear
    For i = 1 To 16
        intComPortExist = EnumSerPorts(i)
        If intComPortExist > 0 Then
            cboPort.AddItem Trim(str(i))
        End If
    Next
    
    cboBaudrate.AddItem ("150")
    cboBaudrate.AddItem ("300")
    cboBaudrate.AddItem ("600")
    cboBaudrate.AddItem ("1200")
    cboBaudrate.AddItem ("2400")
    cboBaudrate.AddItem ("4800")
    cboBaudrate.AddItem ("9600")
    cboBaudrate.AddItem ("14400")
    cboBaudrate.AddItem ("19200")
    cboBaudrate.AddItem ("115200")
    
    cboDatabit.AddItem ("7")
    cboDatabit.AddItem ("8")
    
    cboStartbit.AddItem ("1")
    cboStartbit.AddItem ("2")
    
    cboStopbit.AddItem ("1")
    cboStopbit.AddItem ("1.5")
    cboStopbit.AddItem ("2")
    
    cboParity.AddItem ("N")
    cboParity.AddItem ("E")
    cboParity.AddItem ("O")
    
    txtTCPIP.Text = ""
    txtTCPPort.Text = ""
    
    '==============================
    intPhase = 1
    intBufCnt = 0
    intFrameNo = 1
    intSndPhase = 0
    strState = ""
    blnIsETB = False
    '==============================
    
    If gHOSP.BARUSE = "Y" Then
        optBarSeq(0).Value = True
    Else
        optBarSeq(1).Value = True
    End If
    
    
    cboState.Clear
    cboState.AddItem "--전체--"
    cboState.AddItem "전송"
    cboState.AddItem "미전송"
    cboState.ListIndex = 0
    
    cboRstType.Clear
    cboRstType.AddItem "검사일자"
    cboRstType.AddItem "접수일자"
    cboRstType.ListIndex = 0
    
    txtTV.Text = ""
    lblPatInfo(0).Caption = ""
    lblPatInfo(1).Caption = ""
    lblPatInfo(2).Caption = ""
    
    lblCommStatus.Caption = ""
    
    txtBarcode.Text = ""
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now
    
    fraLinks.Visible = False

End Sub

Private Sub lblActionTest_Click(index As Integer)
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon
    
    If index = 0 Then
        Call GetTestList
    
    ElseIf index = 1 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
            Exit Sub
        End If
        
'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
'            Exit Sub
'        End If
        
        If MsgBox(txtTestNm.Text & "를 삭제하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
             Exit Sub
        End If
        Set Test_Property = New Scripting.Dictionary
    
        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "OCH", txtOChannel.Text
            .Add "TESTCD", txtTestCd.Text
        End With
        
        Set objTest_Property = New clsCommon
        
        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .DelTestInfo(Test_Property) Then
                '-- 삭제 오류
                'Call GetTestList
            End If
        End With
        
        Call GetTestList
        
    ElseIf index = 2 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "검사항목을 먼저 선택하세요", vbCritical, Me.Caption
            Exit Sub
        End If
        
'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
'            txtOChannel.SetFocus
'            Exit Sub
'        End If
'
'        If Trim(txtRChannel.Text) = "" Then
'            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
'            txtRChannel.SetFocus
'            Exit Sub
'        End If
        
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If
        
        Set Test_Property = New Scripting.Dictionary
    
        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            .Add "RES", txtResSpec.Text
            .Add "REFL", txtRefLow.Text
            .Add "REFH", txtRefHigh.Text
            .Add "RSTTYPE", cboResultType.Text
            If optCutUse(0).Value = True Then
                .Add "CUTUSE", "N"
            Else
                .Add "CUTUSE", "Y"
            End If
            .Add "COLIN", txtCOLIn.Text
            .Add "COLCP", cboCOL.Text
            .Add "COLOUT", txtCOLOut.Text
            .Add "COHIN", txtCOHIn.Text
            .Add "COHCP", cboCOH.Text
            .Add "COHOUT", txtCOHOut.Text
            .Add "COMOUT", txtCOMOut.Text
            '-- QC
            .Add "LAB", txtLab.Text
            .Add "LOT", txtLot.Text
            .Add "ANALYTE", txtAnalyte.Text
            .Add "METHOD", txtMethod.Text
            .Add "INSTRUMENT", txtInstrument.Text
            .Add "REAGENT", txtReagent.Text
            .Add "UNIT", txtUnit.Text
            
            '.Add "TEMP", txtTemp.Text
            .Add "TEMP", chkResSpec.Value

        End With
        
        Set objTest_Property = New clsCommon
        
        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With
        
        Call GetTestList
        
    ElseIf index = 3 Then
        If frameOrder.Visible = True Then
            frameOrder.Visible = False
        Else
            frameOrder.Visible = True
        End If
    End If
    
End Sub

Private Sub lblActionTest_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim i As Integer

    For i = 0 To 2
        lblActionTest(i).ForeColor = vbBlack
        shpA(i).BorderColor = &H808080
    Next
    
    lblActionTest(index).ForeColor = vbBlue
    shpA(index).BorderColor = vbCyan


End Sub

Private Sub lblClear_Click()
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0

End Sub

Private Sub lblClear_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblClear.ForeColor = vbBlue
    shpC.BorderColor = vbCyan

End Sub

Private Sub lblComSave_Click()

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If

    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "COMPORT", cboPort.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "SPEED", cboBaudrate.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "PARITY", cboParity.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "DATABIT", cboDatabit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "STARTBIT", cboStartbit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "STOPBIT", cboStopbit.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
    GetSetup
    
    GetCommList
    
    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKCancel, Me.Caption

End Sub

Private Sub lblComSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblComSave.ForeColor = vbBlue
    shpCom.BorderColor = vbCyan

End Sub

Private Sub lblMenu_Click(index As Integer)
    
    frame1.Visible = False
    frame2.Visible = False
    frame3.Visible = False
    'frame4.Visible = False
    fraInterface.Visible = False
    fraResult.Visible = False
    
    lblPatInfo(0).Caption = ""
    lblPatInfo(1).Caption = ""
    lblPatInfo(2).Caption = ""
    lblPatInfo(3).Caption = ""
    lblPatInfo(4).Caption = ""
    
    lblMenu(0).ForeColor = vbBlack
    shpB(0).BorderColor = vbGreen
    
    Select Case index
        Case 0:
                frame1.Visible = True
                frame1.ZOrder 0
        
                fraInterface.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비와 인터페이스]     ◈◈◈◈◈"
                
                tmrComm.Enabled = False
                tmrFlipFlop.Enabled = False
                
                lblCommStatus.Caption = ""
        Case 1:
                frame2.Visible = True
                frame2.ZOrder 0
        
                fraResult.Visible = True
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 결과 조회]     ◈◈◈◈◈"
        Case 2:
                frame3.Visible = True
                frame3.ZOrder 0
    
                '-- 검사코드
                Call GetTestList
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [검사 코드 설정]     ◈◈◈◈◈"
        
        Case 3:
                Frame4.Visible = True
                Frame4.ZOrder 0
    
                '-- 통신설정
                Call GetCommList
                frmMain.Caption = gHOSP.MACHNM & Space$(5) & "◈◈◈◈◈     [장비 통신 설정]     ◈◈◈◈◈"
    
    End Select
    
End Sub

Private Sub lblMenu_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblMenu(i).ForeColor = vbBlack
        shpB(i).BorderColor = vbGreen
    Next
    
    lblMenu(index).ForeColor = vbBlue
    shpB(index).BorderColor = vbCyan

End Sub



Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblSave.ForeColor = vbBlue
    shpS.BorderColor = vbCyan

End Sub

Private Sub lblTcpSave_Click()
    
    If optComType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    If optTCPType(0).Value = True Then
        Call WritePrivateProfileString("COMM", "TCPTYPE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("COMM", "TCPTYPE", "2", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    Call WritePrivateProfileString("COMM", "TCPIP", txtTCPIP.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("COMM", "TCPPORT", txtTCPPort.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
    GetSetup
    
    GetCommList

    MsgBox "통신정보가 변경되었습니다.", vbInformation + vbOKCancel, Me.Caption

End Sub

Private Sub lblTcpSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblTcpSave.ForeColor = vbBlue
    shpTcp.BorderColor = vbCyan

End Sub

Private Sub lblWork_Click()
    
    frmWorkList.Show 'vbModal
    
End Sub

Private Sub lblWork_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblWork.ForeColor = vbBlue
    shpW.BorderColor = vbCyan

End Sub

Private Sub optComType_Click(index As Integer)
    
    If index = 0 Then
        frameCom.Enabled = True
        frameTCP.Enabled = False
    Else
        frameCom.Enabled = False
        frameTCP.Enabled = True
    End If

End Sub

Private Sub optCutUse_Click(index As Integer)
    If index = 0 Then
        frameCutOff.Enabled = False
    Else
        frameCutOff.Enabled = True
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer

    For i = 0 To 3
        lblMenu(i).ForeColor = vbBlack
        shpB(i).BorderColor = vbGreen
    Next
    
    lblWork.ForeColor = vbBlack
    lblSave.ForeColor = vbBlack
    lblClear.ForeColor = vbBlack
    shpW.BorderColor = &H808080
    shpS.BorderColor = &H808080
    shpC.BorderColor = &H808080
    
    lblResult.ForeColor = vbBlack
    shpR.BorderColor = &H808080
    
    
End Sub



Private Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol As Integer
    
    '-- 정렬
    If Row = 0 Then
        '-- 정렬 추가
        
        Exit Sub
    End If
    
    '-- 환자정보표시
    
    '-- 결과표시
    If GetPatTRestResult(Row) = -1 Then
        '장비결과가 없을경우 검사명만 보여주기
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '◆
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 12
                End If
            Next
        End With
    End If
        
End Sub

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Private Function GetPatTRestResult(ByVal asRow As Integer) As Integer
    Dim strBarno As String
    Dim intSeq   As String
    Dim strExamDate As String
    Dim intRow   As Integer
    
On Error GoTo RST

    GetPatTRestResult = -1
    intRow = 0
    
    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = Mid(GetText(spdOrder, asRow, colEXAMDATE), 1, 8)
    
    If intSeq = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EXAMNAME, RESULT" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
'    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                AdoRs_Local.MoveNext
            Loop
        End With
        GetPatTRestResult = 1
    End If
    
    AdoRs_Local.Close
    
Exit Function

RST:
    GetPatTRestResult = -1

End Function

'인터페이스 환자선택시 우측에 검사항목/결과보여주기
Public Function GetPatTRestResult_Search(ByVal asRow As Integer) As Integer
    Dim strBarno As String
    Dim intSeq   As String
    Dim strExamDate As String
    Dim intRow   As Integer
    
On Error GoTo RST

    GetPatTRestResult_Search = -1
    intRow = 0
    
    intSeq = GetText(spdROrder, asRow, colSAVESEQ)
    strExamDate = Mid(GetText(spdROrder, asRow, colEXAMDATE), 1, 8)
    
    If intSeq = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO,EXAMCODE,EQUIPCODE,EXAMNAME,EQUIPRESULT,RESULT,REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
'    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdRResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colRSEQNO)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                Call SetText(frmMain.spdRResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                If AdoRs_Local.Fields("REFJUDGE").Value & "" <> "" Then
                    SetForeColor frmMain.spdRResult, intRow, intRow, colRLISRESULT, colRLISRESULT, 255, 0, 0
                End If
                
                If AdoRs_Local.Fields("EXAMCODE").Value & "" = "24HRS-V" Then
                    txtTV.Text = AdoRs_Local.Fields("RESULT").Value & ""
                End If
                AdoRs_Local.MoveNext
            Loop
        End With
        GetPatTRestResult_Search = 1
    End If
    
    AdoRs_Local.Close
    
Exit Function

RST:
    GetPatTRestResult_Search = -1

End Function


Private Sub spdOrdMst_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
        
    If KeyAscii = vbKeyReturn Then
        '-- Delete
              SQL = ""
        SQL = SQL & "DELETE FROM ORDMASTER "
        
        Call DBExec(AdoCn_Local, SQL)
        
        'Insert
        For intRow = 1 To spdOrdMst.MaxRows
                  SQL = ""
            SQL = SQL & "INSERT INTO ORDMASTER (ORDERCODE,ORDERNAME) VALUES ("
            SQL = SQL & "'" & GetText(spdOrdMst, intRow, 1) & "','')"
            
            Call DBExec(AdoCn_Local, SQL)
        Next
    End If
    
End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Call SetSpreadSort(spdTest, 0)
        Exit Sub
    End If
    
    With spdTest
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
        txtRefLow.Text = GetText(spdTest, Row, colLLOW)
        txtRefHigh.Text = GetText(spdTest, Row, colLHIGH)
        cboResultType.Text = GetText(spdTest, Row, colLRSTTYPE)
        If GetText(spdTest, Row, colLCUTUSE) = "1" Then
            optCutUse(1).Value = True
        Else
            optCutUse(0).Value = True
        End If
        txtCOLIn.Text = GetText(spdTest, Row, colLCOLIN)
        cboCOL.Text = GetText(spdTest, Row, colLCOLCOMP)
        txtCOLOut = GetText(spdTest, Row, colLCOLOUT)
        txtCOHIn.Text = GetText(spdTest, Row, colLCOHIN)
        cboCOH.Text = GetText(spdTest, Row, colLCOHCOMP)
        txtCOHOut = GetText(spdTest, Row, colLCOHOUT)
        txtCOMOut = GetText(spdTest, Row, colLCOMOUT)
        '-- QC
        txtLab = GetText(spdTest, Row, colLQCLab)
        txtLot = GetText(spdTest, Row, colLQCLot)
        txtAnalyte = GetText(spdTest, Row, colLQCAnalyte)
        txtMethod = GetText(spdTest, Row, colLQCMethod)
        txtInstrument = GetText(spdTest, Row, colLQCInstrument)
        txtReagent = GetText(spdTest, Row, colLQCReagent)
        txtUnit = GetText(spdTest, Row, colLQCUnit)
        txtTemp = GetText(spdTest, Row, colLQCTemp)
        
        If GetText(spdTest, Row, colLUseResSpec) = "1" Then
            chkResSpec.Value = "1"
        Else
            chkResSpec.Value = "0"
        End If
    
    End With
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Len(Trim(txtBarcode)) = 9 And Trim(txtBarcode) <> "" And KeyCode = vbKeyReturn Then
        txtBarcode.Text = Mid(txtBarcode.Text, 1, 2) & "-" & Mid(Trim(txtBarcode), 3) & "-" & "1"
        mResult.BarNo = Trim(txtBarcode.Text)
        Call SetPatInfo_WithBar(mResult.BarNo, gHOSP.RSTTYPE)
        txtBarcode.SelStart = 0
        txtBarcode.SelLength = Len(txtBarcode.Text)
    End If
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo_WithBar(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) Then
                If Val(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) = Val(pBarno) Then
                    If Trim(GetText(frmMain.spdOrder, i, colSTATE)) = "" Then
                        intRow = i
                        Exit For
                    End If
                End If
            End If
        Next i
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
        'Call SetText(.spdOrder, Format(frmMain.txtSeqNo.Text, "#0"), intRow, colSEQNO)
        
        'SetText SPD, Format(frmMain.txtSeqNo.Text, "#0"), asRow, colSEQNO
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub


Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    Dim strTestNm   As String
    
    If KeyAscii = vbKeyReturn Then
        If Trim(txtTestCd.Text) <> "" Then
            strTestNm = GetTest(Trim(txtTestCd.Text))
            If strTestNm <> "" Then
                txtTestNm = strTestNm
                txtAbbrNm = strTestNm
            End If
        End If
    End If
End Sub

Private Sub txtTV_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call cmdTVSave_Click
    End If
    
End Sub

Private Sub wsck_ConnectionRequest(ByVal requestID As Long)

    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.accept requestID
        lblStatus.Caption = "장비에 접속되었습니다."
    End If

End Sub

Private Sub wsck_DataArrival(ByVal bytesTotal As Long)
    Dim strText As String
    Dim strTmp As String
    
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strSendAck  As String
    Dim strRcvCnt   As String
    
    Dim strNS       As String
    Dim strNE       As String
    Dim intNS       As Integer
    Dim intNE       As Integer
    
    Dim strSendData  As String
    Dim varBuffers   As Variant
    Dim i As Integer
    Dim lngBufLen As Long
    Dim BufChar     As String
    
    wSck.GetData strText

    pBuffer = strText
    
    dtpToday.Value = Now
    
    SetRawData "[Rx]" & pBuffer
    
    Call TCP_Protocol
    
    tmrResult.Interval = 1000
    tmrResult.Enabled = True
    
End Sub


