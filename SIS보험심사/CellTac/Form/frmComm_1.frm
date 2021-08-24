VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11985
   WindowState     =   2  '최대화
   Begin VB.PictureBox picWork 
      Height          =   5910
      Left            =   30
      ScaleHeight     =   5850
      ScaleWidth      =   11865
      TabIndex        =   12
      Top             =   570
      Width           =   11925
      Begin TabDlg.SSTab tabWork 
         Height          =   5850
         Left            =   45
         TabIndex        =   13
         Top             =   0
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   10319
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   " Work List"
         TabPicture(0)   =   "frmComm_1.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdLoad"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lvwWorkList"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtWorkSpcno"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   " 저장된 결과 "
         TabPicture(1)   =   "frmComm_1.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "pnlCom"
         Tab(1).Control(1)=   "lvwComplete"
         Tab(1).Control(2)=   "lvwCuData"
         Tab(1).Control(3)=   "pnlCom2"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   " 미 저장 결과 "
         TabPicture(2)   =   "frmComm_1.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label6"
         Tab(2).Control(1)=   "Label5"
         Tab(2).Control(2)=   "Label10"
         Tab(2).Control(3)=   "cmdOk"
         Tab(2).Control(4)=   "lvwError"
         Tab(2).Control(5)=   "txtEqpno"
         Tab(2).Control(6)=   "txtSpcno"
         Tab(2).ControlCount=   7
         Begin VB.TextBox txtWorkSpcno 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            Height          =   270
            Left            =   1140
            TabIndex        =   0
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txtSpcno 
            Appearance      =   0  '평면
            Height          =   270
            Left            =   -69585
            TabIndex        =   16
            Top             =   435
            Width           =   1575
         End
         Begin VB.TextBox txtEqpno 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Height          =   270
            Left            =   -72255
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   435
            Width           =   1560
         End
         Begin MSComctlLib.ListView lvwError 
            Height          =   4995
            Left            =   -74955
            TabIndex        =   14
            Top             =   780
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   8811
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin HSCotrol.CButton cmdOk 
            Height          =   330
            Left            =   -64350
            TabIndex        =   17
            Top             =   405
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            Caption         =   "확 인"
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
         Begin HSCotrol.UserPanel pnlCom2 
            Height          =   5385
            Left            =   -68700
            TabIndex        =   21
            Top             =   540
            Visible         =   0   'False
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   9499
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtCOM2 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   315
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   30
               Top             =   405
               Width           =   5730
            End
            Begin VB.Frame Frame2 
               Height          =   645
               Left            =   60
               TabIndex        =   22
               Top             =   4650
               Width           =   5760
               Begin MSComDlg.CommonDialog cdlFile 
                  Left            =   5265
                  Top             =   60
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin HSCotrol.CButton cmdChksum 
                  Height          =   360
                  Left            =   2205
                  TabIndex        =   23
                  Top             =   180
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   635
                  Caption         =   "SUM"
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
                  BorderColor     =   8421504
               End
               Begin HSCotrol.CButton cmdCOMOutput2 
                  Height          =   360
                  Left            =   1155
                  TabIndex        =   24
                  Top             =   180
                  Width           =   1000
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Send"
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
                  BorderColor     =   8421504
               End
               Begin HSCotrol.CButton cmdCOMClear2 
                  Height          =   360
                  Left            =   3600
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Clear"
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
               Begin HSCotrol.CButton cmdCOMInput2 
                  Height          =   360
                  Left            =   90
                  TabIndex        =   26
                  Top             =   180
                  Width           =   1000
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Receive"
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
                  BorderColor     =   8421504
               End
               Begin HSCotrol.CButton cmdCOMLoad 
                  Height          =   360
                  Left            =   4665
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "File Load"
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
               Begin HSCotrol.CButton cmdACK 
                  Height          =   360
                  Left            =   3105
                  TabIndex        =   28
                  Top             =   180
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   635
                  Caption         =   "ACK"
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
                  BorderColor     =   8421504
               End
               Begin HSCotrol.CButton cmdENQ 
                  Height          =   360
                  Left            =   2655
                  TabIndex        =   29
                  Top             =   180
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   635
                  Caption         =   "ENQ"
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
                  BorderColor     =   8421504
               End
            End
         End
         Begin MSComctlLib.ListView lvwCuData 
            Height          =   5415
            Left            =   -69195
            TabIndex        =   31
            Top             =   585
            Visible         =   0   'False
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   9551
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvwComplete 
            Height          =   5400
            Left            =   -74955
            TabIndex        =   32
            Top             =   405
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   9525
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin HSCotrol.UserPanel pnlCom 
            Height          =   5400
            Left            =   -74955
            TabIndex        =   33
            Top             =   315
            Visible         =   0   'False
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   9525
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.Frame Frame1 
               Height          =   645
               Left            =   45
               TabIndex        =   35
               Top             =   4650
               Width           =   11610
               Begin HSCotrol.CButton cmdCOMSave 
                  Height          =   360
                  Left            =   10515
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "File Save"
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
               Begin HSCotrol.CButton cmdCOMOutput 
                  Height          =   360
                  Left            =   1155
                  TabIndex        =   37
                  Top             =   180
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Send"
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
                  BorderColor     =   8421504
               End
               Begin HSCotrol.CButton cmdCOMClear 
                  Height          =   360
                  Left            =   9450
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Clear"
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
               Begin HSCotrol.CButton cmdCOMInput 
                  Height          =   360
                  Left            =   90
                  TabIndex        =   39
                  Top             =   180
                  Width           =   1000
                  _ExtentX        =   1773
                  _ExtentY        =   635
                  Caption         =   "Receive"
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
                  BorderColor     =   8421504
               End
            End
            Begin VB.TextBox txtCom 
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4395
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   34
               Top             =   270
               Width           =   11595
            End
         End
         Begin MSComctlLib.ListView lvwWorkList 
            Height          =   5040
            Left            =   30
            TabIndex        =   40
            Top             =   780
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   8890
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin HSCotrol.CButton cmdLoad 
            Height          =   330
            Left            =   2460
            TabIndex        =   41
            Top             =   420
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            Caption         =   "Load"
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "검체번호"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   210
            TabIndex        =   42
            Top             =   480
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "검체 번호 연결 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   -74835
            TabIndex        =   20
            Top             =   480
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "검체 번호 :"
            Height          =   180
            Left            =   -70530
            TabIndex        =   19
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "장비 번호 :"
            Height          =   180
            Left            =   -73185
            TabIndex        =   18
            Top             =   480
            Width           =   900
         End
      End
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4695
      Top             =   6510
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5175
      Top             =   6510
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   2955
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":0054
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":05EE
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":0B88
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":1122
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":16BC
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":1C56
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   3795
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5640
      Top             =   6510
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
            Picture         =   "frmComm_1.frx":21F0
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":278A
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":2D24
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":32BE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3B50
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3CAA
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3E04
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   2
      Top             =   6495
      Width           =   11940
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6375
         TabIndex        =   3
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Run"
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   7740
         TabIndex        =   4
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Stop"
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   9120
         TabIndex        =   5
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
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
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   10485
         TabIndex        =   6
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
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
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "작업대기 중.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   960
         TabIndex        =   11
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " 상태 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   210
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm_1.frx":3F5E
      Caption         =   " Communication"
      SubCaption      =   "검사 장비와 통신하여 결과를 저장 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Left            =   10140
         TabIndex        =   9
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   9105
         TabIndex        =   8
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   8040
         TabIndex        =   7
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11010
         Picture         =   "frmComm_1.frx":51E0
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_1.frx":576A
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_1.frx":5CF4
         Top             =   255
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "순서"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "등록번호"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "성  명"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "검체번호"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "검체번호"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "상 태"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "검사항목"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Public WithEvents WorkList As clsMsg_Result
Attribute WorkList.VB_VarHelpID = -1
Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult

Private mAdoRs      As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer As String

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String
Dim fCellTac(100) As String
Dim fCellItem(20) As String
Dim fCellItemNm(20) As String

Private Function f_funGet_JobList(ByVal strKeyno As String) As String

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strOrder    As String
    
    sqlDoc = "select TESTCD from LIMAS301 where SPCNO = " & STS(strKeyno)
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "TESTCD = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + "0"
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    f_funGet_JobList = strOrder
End Function

Public Function p_funGet_Bn2BCC_2(ByVal strPara As String) As Integer

    Dim intIdx  As Integer
    Dim intBcc  As Integer
    
    intBcc = 0
    For intIdx = 1 To Len(strPara)
        intBcc = intBcc + Asc(Mid$(strPara, intIdx, 1))
    Next
    
    p_funGet_Bn2BCC_2 = 0
    If intBcc Mod 64 = 0 Then p_funGet_Bn2BCC_2 = 1

End Function


Private Sub Set_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub

Private Sub SetListHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
        End With
        .HideColumnHeaders = False
    End With
    
    'Work List
    With lvwWorkList
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        Call SetlistView_Complete(lvwWorkList)
    End With
    
    '검사 완료 테이블
    With lvwComplete
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        Call SetlistView_Complete(lvwComplete)
    End With
    
    '미등록 테이블
    With lvwError
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        Call SetlistView_Complete(lvwError)
    End With
    
End Sub

Private Sub SetlistView_Complete(lvw As Listview)
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    Dim ii              As Integer
    
    Set objHeadeItem = New clsCommon
    With objHeadeItem
        .SetAdoCn AdoCn_SQL
        Set mAdoRs = .Get_TestItemList(INS_CODE)
    End With
    Set objHeadeItem = Nothing
    
    Erase fCellItem
    Erase fCellItemNm
    
    If Not mAdoRs Is Nothing Then
        With lvw
            .ColumnHeaders.Clear
            Call .ColumnHeaders.Add(, "EQP_ID", "검체 번호")
            Do Until mAdoRs.EOF
                ii = ii + 1
                Set itemH = .ColumnHeaders.Add
                With itemH
                    '컬럽 헤더키를 장비검사 코드로
                    .Key = COL_KEY & Trim(mAdoRs.Fields("ITEMCODE") & "")
                    fCellItem(ii) = COL_KEY & Trim(mAdoRs.Fields("ITEMCODE") & "")
                    '컬럽명은 검사 항목 이름
                    .Text = Trim(mAdoRs.Fields("INFORMALNAME") & "")
                    'fCellItemNm(ii) = ii
                    '테그는 검사 코드로
                    .Tag = Trim(mAdoRs.Fields("ITEMCODE") & "")
                End With
                Set itemH = Nothing
                mAdoRs.MoveNext
            Loop
        End With
    End If
    Set mAdoRs = Nothing
End Sub

Private Sub SetItem_List()

'    Dim objItem As clsCommon
'    Dim itemX As ListItem
'
'On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Private Sub SetItem_List()"
'    Set objItem = New clsCommon
'    '검사 코드 메치테이블
'    With objItem
'        .SetAdoCn AdoCn_Jet
'        Set mAdoRs = .Get_TestItemList(INS_CODE)
'    End With
'
'    lvwCuData.ListItems.Clear
'    If Not mAdoRs Is Nothing Then
'        Do Until mAdoRs.EOF
'            Set itemX = lvwCuData.ListItems.Add(, , Trim(mAdoRs.Fields("TESTCD_EQP") & ""), , "LST")
'                itemX.SubItems(1) = Trim(mAdoRs.Fields("TESTCD") & "")
'                itemX.SubItems(2) = Trim(mAdoRs.Fields("TESTNM") & "")
'                itemX.SubItems(3) = ""
'                itemX.Tag = Trim(mAdoRs.Fields("TESTCD_EQP") & "")
'            Set itemX = Nothing
'            mAdoRs.MoveNext
'        Loop
'    End If
'
'    Set objItem = Nothing
'    Set mAdoRs = Nothing
'
'Exit Sub
'ErrRoutine:
'    Set objItem = Nothing
'    Set mAdoRs = Nothing
'    Call ErrMsgProc(CallForm)

    Dim objItem As clsCommon
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub SetItem_List()"
    Set objItem = New clsCommon
    '검사 코드 메치테이블
    With objItem
        '.SetAdoCn AdoCn_Jet
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_TestItemList(INS_CODE)
    End With
    
    lvwCuData.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        Do Until mAdoRs.EOF
'            Set itemX = lvwCuData.ListItems.Add(, , Trim(mAdoRs.Fields("ITEMCODE") & ""), , "LST")
            Set itemX = lvwCuData.ListItems.Add(, , Trim(INS_CODE), , "LST")
                itemX.SubItems(1) = Trim(mAdoRs.Fields("INFORMALNAME") & "")
                itemX.SubItems(2) = Trim(mAdoRs.Fields("ITEMCODE") & "")
                itemX.SubItems(3) = ""
'                itemX.Tag = Trim(mAdoRs.Fields("ITEMCODE") & "")
                itemX.Tag = Trim(INS_CODE)
            Set itemX = Nothing
            mAdoRs.MoveNext
        Loop
    End If
    
    Set objItem = Nothing
    Set mAdoRs = Nothing
    Exit Sub
ErrRoutine:
    Set objItem = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdACK_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ACK))

End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call cmdRun_Click
        Case 1
            Call cmdStop_Click
        Case 2
            Call cmdClear_Click
        Case 3 'cmd close
            Call cmdClose_Click
        Case Else
    End Select

End Sub

Private Sub cmdClear_Click()
    
    Dim itemX As ListItem
    
    For Each itemX In lvwCuData.ListItems
        itemX.SubItems(3) = ""
    Next
    
    lvwWorkList.ListItems.Clear
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun_Click()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    
        'Call comEQP_OnComm
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
        
    End If
    
    tabWork.Tab = 0
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop_Click()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun_Click()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("중지 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdENQ_Click()
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))
End Sub

Private Sub cmdLoad_Click()
    Dim itemH           As ColumnHeader
    Dim objWorkList    As clsCommon
    Dim itemX As ListItem
    
    Set objWorkList = New clsCommon
    With objWorkList
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_WorkList(Format(txtWorkSpcno.Text, "yyyymmdd"))
    
    End With
    
    Set objWorkList = Nothing
    
    If Not mAdoRs Is Nothing Then
        lvwWorkList.ListItems.Clear
        Do Until mAdoRs.EOF
            Set itemX = lvwWorkList.ListItems.Add(, , Trim(mAdoRs.Fields("SPCNO")))
            If Not itemX Is Nothing Then
                With itemX
                    .Key = COL_KEY & Trim(mAdoRs.Fields("SPCNO"))
                    .Tag = Trim(mAdoRs.Fields("SPCNO"))
                    .SmallIcon = "ITM"
                End With
            End If
            mAdoRs.MoveNext
        Loop
    End If
    
    Set mAdoRs = Nothing

End Sub

Private Sub cmdOk_Click()
'    Dim itemX       As ListItem
'    Dim itemA       As ListItem
'    Dim itemS       As ListSubItem
'    Dim objSave     As clsEqpResult
'
'    CallForm = "frmComm - Private Sub cmdOk_Click()"
'On Error GoTo ErrorRoutine
'    '미등록 검사 결과에 검체 번호 메치
'    If Trim(txtSpcno) = "" Then
'        Call ShowMessage("검체 번호를 입력 하세요.")
'        txtSpcno.SetFocus
'        Exit Sub
'    End If
'
'    Set itemX = lvwError.SelectedItem
'    If Not itemX Is Nothing Then
'        Set objSave = New clsEqpResult
'        With objSave
'            If .Spc_Exists(Trim(txtSpcno)) Then '검체번호 가 있는지 검사
'                .EQPNUM = itemX.Text
'                .SPCID = Trim(txtSpcno)
'                .SPCTYPE = itemX.Tag
'                For Each itemS In itemX.ListSubItems
'                    '서브아이템에 검사 결과 가 있으면
'                    If Trim(itemS.Text) <> "" Then
'                        Call .Set_EqpResultjet(lvwError.ColumnHeaders(itemS.Index + 1).Tag, itemS.Text)
'                        Call .Set_EqpResultsql(lvwError.ColumnHeaders(itemS.Index + 1).Tag, itemS.Text, itemS.Tag)
'                    End If
'                Next
'                '검사결과 등록 테이블에 등록
'                Set itemA = lvwComplete.FindItem(Trim(txtSpcno), lvwText, , lvwWhole)
'                If itemA Is Nothing Then
'                    Set itemA = lvwComplete.ListItems.Add()
'                End If
'                With itemA
'                    .Key = COL_KEY & Trim(txtSpcno)
'                    .Text = Trim(txtSpcno)
'                    .Tag = itemX.Tag
'                    .SmallIcon = "LST"
'                    For Each itemS In itemX.ListSubItems
'                        .SubItems(itemS.Index) = itemS.Text
'
'                    Next
'                End With
'                '미등록 테이블에서 제거
'                Call lvwError.ListItems.Remove(itemX.Index)
'            Else
'                Call ShowMessage("지정한 검체 번호가 발행 되지 않았습니다.")
'            End If
'        End With
'
'        Set itemX = Nothing
'        Set itemA = Nothing
'        Set itemS = Nothing
'
'        Set objSave = Nothing
'        txtEqpno = ""
'        txtSpcno = ""
'    End If
'Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'    Set itemA = Nothing
'    Set itemS = Nothing
'    Set objSave = Nothing
'    Call ErrMsgProc(CallForm)
End Sub

Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    
    Select Case comEQP.CommEvent
        Case comEvReceive
        
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            
            Arr = comEQP.Input
            'Debug.Print Arr
            Call ComReceive(Arr)
            
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) 변경 감지"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) 변경 감지"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) 변경 감지"
        Case comEvRing
            strEVMsg = " 전화 벨이 울리는 중"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) 감지"

        ' 오류 메시지
        Case comBreak
            strERMsg = " 중단 신호 수신"
        Case comCDTO
            strERMsg = " 반송파 검출 시간 초과"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) 시간 초과"
        Case comDCB
            strERMsg = " 포트에 대한 장치 제어 블록(DCB) 검색 중 예기치 못한 오류"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) 시간 초과"
        Case comFrame
            strERMsg = " 프레이밍 오류"
        Case comOverrun
            strERMsg = " 패리티 오류"
        Case comRxOver
            strERMsg = " 수신 버퍼 초과"
        Case comRxParity
            strERMsg = " 패리티 오류"
        Case comTxFull
            strERMsg = " 전송 버퍼에 여유가 없음"
        Case Else
            strERMsg = " 알 수 없는 오류 또는 이벤트"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
End Sub

Private Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Private Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
Dim Pos_point As Integer
    Do
        Pos_point = InStr(FSend_Str, FCheck_Char)
        If Pos_point < 1 Then
            Exit Do
        ElseIf Pos_point = 1 Then
            FSend_Str = FChange_Char + Mid$(FSend_Str, 2)
        Else
            FSend_Str = Mid$(FSend_Str, 1, Pos_point - 1) + FChange_Char + Mid$(FSend_Str, Pos_point + 1)
        End If
    Loop
    Text_Change = FSend_Str
End Function


Private Sub ComReceive(ByRef RecData() As Byte)
    
    Dim strTemp     As String
    Dim strRec      As String
    Dim intEOT_POS  As Long
    
    Dim sStxCheck As Integer
    Dim sEtxCheck As Integer
    Dim com_sTemp As String
    
    Dim intIdx  As Integer, strSend As String
    Dim ii      As Integer
    
    Dim Channel_No  As String       ' 문자형 변수
    Dim pGrid_Point As Integer      ' 해당 검사자 Point
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수

    Dim pDoCount    As Integer
    Dim Loop_Count  As Integer
    Dim sChannel As String, sRstText As String
    Dim itemX       As ListItem
    Dim itemC       As ListItem
    Dim tmpRst      As String
    
    Static OrgMsg As String
    'strRec = StrConv(RecData, vbUnicode)
    strRec = RecData
'    Debug.Print strRec
    'f_strBuffer = ""
    Call COM_INPUT(strRec)
    
    For intIdx = 1 To Len(strRec)
        f_strBuffer = f_strBuffer + Mid$(strRec, intIdx, 1)
        sStxCheck = InStr(f_strBuffer, Chr(2))
        sEtxCheck = InStr(f_strBuffer, Chr(3))
        If sStxCheck <> 0 And sEtxCheck <> 0 Then
            'com_sTemp = Mid$(f_strBuffer, sStxCheck + 1, sEtxCheck - 2)
            com_sTemp = Mid$(f_strBuffer, sStxCheck + 1, sEtxCheck - 2)
            sRstText = f_strBuffer
            '------------------------------<<< fCellTac() 배열 Clear 한다.         >>>----------
            For Loop_Count = 1 To 100: fCellTac(Loop_Count) = "": Next Loop_Count
            '------------------------------<<< fCellTac() 배열에 구분하여 넣는다.  >>>----------
                    
            Do While InStr(f_strBuffer, Chr$(13)) > 0
                pDoCount = pDoCount + 1
                
                fCellTac(pDoCount) = Text_Redefine(sRstText, Chr$(13))
                
                fCellTac(pDoCount) = Replace(fCellTac(pDoCount), "*", "")
                fCellTac(pDoCount) = Replace(fCellTac(pDoCount), "H", "")
                fCellTac(pDoCount) = Replace(fCellTac(pDoCount), "L", "")
                fCellTac(pDoCount) = Replace(fCellTac(pDoCount), "h", "")
                fCellTac(pDoCount) = Replace(fCellTac(pDoCount), "l", "")
                
                'If InStr(fCellTac(pDoCount), "*") > 0 Or InStr(fCellTac(pDoCount), "*") > L Or InStr(fCellTac(pDoCount), "*") > 0 Then
                
                sRstText = Mid$(sRstText, InStr(sRstText, Chr$(13)) + 1)   ' 구분자가 "Chr$(13)" 이다....
                If Len(Trim(sRstText)) = 0 Then Exit Do
                If pDoCount > 99 Then
                    sRstText = ""
                    Exit Do
                End If
                'pGrid_Point = 1                             ' Case by Case 작업일때는 1로 Setting
                Max_Arary_Cnt = lvwWorkList.ListItems.Count              ' Case by Case 작업일때는 Row 값이 총 검사건수
                fCellTac(pDoCount) = Text_Change(fCellTac(pDoCount), Chr$(10), "")
            Loop
                
                
            Set itemX = lvwWorkList.FindItem(Trim(fCellTac(2)), lvwTag, , lvwWhole)
            
            If Not itemX Is Nothing Then
                If Not itemX Is Nothing Then
                    With itemX
                        For ii = 1 To lvwWorkList.ColumnHeaders.Count - 1
                            Select Case ii
                            Case 1, 5, 6, 7: itemX.SubItems(ii + 1) = Trim(fCellTac(ii + 3))
                            Case 2: itemX.SubItems(ii - 1) = Trim(fCellTac(ii + 3))
                            Case 8: itemX.SubItems(ii - 3) = Trim(fCellTac(ii + 3))
                            Case Else: itemX.SubItems(ii) = Trim(fCellTac(ii + 3))
                            End Select
                            
                            Set Result = New clsMsg_Result
                            
                            With Result
                                .MSG_SID = fCellTac(2)
                                .MSG_TYPE = "G"
                                Select Case ii
                                Case 1: .MSG_RESULT = Trim(fCellTac(ii + 4))
                                Case 5: .MSG_RESULT = Trim(fCellTac(ii + 6))
                                Case 2, 6, 7, 8: .MSG_RESULT = Trim(fCellTac(ii + 2))
                                Case Else: .MSG_RESULT = Trim(fCellTac(ii + 3))
                                End Select
                                '.MSG_RESULT = Trim(fCellTac(ii + 3))
                            End With
                            Set Result = Nothing
                            
                            Set Result1 = New clsResult
                            With Result1
                                .Rst_Sid = Trim(fCellTac(2))
                                .Rst_Eid = INS_CODE
                                .Rst_Type = "G"
                                .Rst_Test = fCellItem(ii)
                                Select Case ii
                                Case 1: .Rst_Values = Trim(fCellTac(ii + 4))
                                Case 5: .Rst_Values = Trim(fCellTac(ii + 6))
                                Case 2, 6, 7, 8: .Rst_Values = Trim(fCellTac(ii + 2))
                                Case Else: .Rst_Values = Trim(fCellTac(ii + 3))
                                End Select
                                '.Rst_Values = Trim(fCellTac(ii + 3))
                                .Rst_Tag = Trim(fCellTac(2))
                                .Rst_Error = ""
                            End With
                            
                            Call Result_MsgSplit(Result1)
                        Next
                    End With
                End If
            End If

            Set Result1 = Nothing
                    
            f_strBuffer = ""
        End If
    Next
    
    Exit Sub
    
    For intIdx = 1 To Len(strRec)
        Select Case Mid$(strRec, intIdx, 1)
            Case Chr(2) '-- STX
                        f_strBuffer = Mid$(strRec, intIdx, 1)
            Case Chr(3) '-- ETX
                        f_strBuffer = f_strBuffer + Mid$(strRec, intIdx, 1)
                        Select Case Mid$(f_strBuffer, 2, 1)
                            Case "J"    '-- job list 요청시 처리
                                If Len(f_strBuffer) > 10 Then
                                    strTemp = Mid$(f_strBuffer, 2)
                                    If p_funGet_Bn2BCC_1(Mid$(strTemp, 1, Len(strTemp) - 3)) = Mid$(f_strBuffer, Len(f_strBuffer) - 3, 1) Then
                                        strSend = Mid$(f_strBuffer, 1, 31) + String(12, " ") + "0"
                                        strSend = strSend + f_funGet_JobList(Trim$(Mid$(f_strBuffer, 3, 29)))
                                        strSend = strSend + p_funGet_Bn2BCC_1(Mid$(strSend, 2)) + Chr(13) + Chr(3) + Chr(6)
                                        
                                        Call COM_OUTPUT(strSend)
                                    Else
                                        Call COM_OUTPUT(Chr(21))
                                    End If
                                End If
                            Case "D"    '-- 검사결과 받았을 때 처리
                                If Len(f_strBuffer) > 10 Then
                                    strTemp = Mid$(f_strBuffer, 2)
                                    If p_funGet_Bn2BCC_1(Mid$(strTemp, 1, Len(strTemp) - 3)) = Mid$(f_strBuffer, Len(f_strBuffer) - 3, 1) Then
                                        Set Result = New clsMsg_Result
                                        With Result
                                            .MSG_SID = Trim$(Mid$(f_strBuffer, 2, 29))
                                            .MSG_TYPE = "G"
                                            .MSG_RESULT = f_strBuffer
                                        End With
                                        Set Result = Nothing
                                        
                                        Set Result1 = New clsResult
                                        With Result1
                                            .Rst_Sid = Trim$(Mid$(f_strBuffer, 3, 29))
                                            .Rst_Eid = Trim$(Mid$(f_strBuffer, 3, 29))
                                            .Rst_Type = "G"
                                            .Rst_Test = Mid$(f_strBuffer, 32, 2)
                                            .Rst_Values = Trim$(Mid$(f_strBuffer, 38, 15))
                                            .Rst_Tag = Mid$(f_strBuffer, 32, 2)
                                            .Rst_Error = ""
                                        End With
                                        Call Result_MsgBegin(Trim$(Mid$(f_strBuffer, 3, 29)))
                                        Call Result_MsgSplit(Result1)
                                        Set Result1 = Nothing
                                        
                                        Call COM_OUTPUT(Chr(6))
                                    Else
                                        Call COM_OUTPUT(Chr(21))
                                    End If
                                End If
                        End Select
                        
            Case Chr(21) '-- NAK
                        Call COM_OUTPUT(strSend)
            Case Chr(6) '-- ACK
                        strSend = ""
            Case Else:  f_strBuffer = f_strBuffer + Mid$(strRec, intIdx, 1)
                
        End Select
    Next

'-- YEJ
''Lf가 있으면 ACK
'    If InStr(1, strRec, MSG_LF) > 0 Then
'        Call COM_OUTPUT(MSG_ACK)
'    End If
''Enq가 있으면 ACK
'    If InStr(1, strRec, MSG_ENQ) > 0 Then
'        Call COM_OUTPUT(MSG_ACK)
'    End If
''ACK가 있고 현제 오더전송 중
'    If (InStr(1, strRec, MSG_ACK) > 0) And (CU_STATUS = MSG_Q) Then
'        Call Order_Ready(Chr(COM_ACK))
'    End If
''NACK가 있고 현제 오더전송 중
'    If (InStr(1, strRec, MSG_NAK) > 0) And (CU_STATUS = MSG_Q) Then
'        Call Order_Ready(Chr(COM_NACK))
'    End If
''EOT가 있으면 EOT
'    If (InStr(1, strRec, MSG_EOT) > 0) Then
'        If (CU_STATUS = MSG_Q) Then
'            Call Order_Ready(MSG_EOT)
'            Call Replace(strRec, MSG_EOT, "")
'        Else
'            Call COM_OUTPUT(MSG_EOT)
'        End If
'    End If
'
''문자열에 ACK 가 있으면 "" 로 대체
'    Call Replace(strRec, MSG_ACK, "")
'    OrgMsg = OrgMsg & strRec
'
'    strTemp = ""
'    'EOT 가 있는지 검사
'    intEOT_POS = InStr(1, OrgMsg, MSG_EOT)
'    If intEOT_POS > 0 Then
'        Debug.Print Mid(OrgMsg, 1, intEOT_POS)
'        'EOT가 있으면 EOT까지의 문자열을 분석
'        Call COM_INPUT_String(Mid(OrgMsg, 1, intEOT_POS))
'        Call Msg_Analysis(Mid(OrgMsg, 1, intEOT_POS))
'        'EOT이후 문자열은 다음 분석을 위해 저장
'        strTemp = Mid(OrgMsg, intEOT_POS + 1)
'    Else
'        strTemp = OrgMsg
'    End If
'
'    OrgMsg = strTemp

End Sub

Public Function p_funGet_Bn2BCC_1(ByVal strPara As String) As String

    Dim intIdx  As Integer
    Dim intBcc  As Integer
    
    intBcc = 0
    For intIdx = 1 To Len(strPara)
        intBcc = intBcc + Asc(Mid$(strPara, intIdx, 1))
    Next
    
    intBcc = 64 - (intBcc Mod 64)
    
    If intBcc < 32 Then intBcc = intBcc + 64
    
    p_funGet_Bn2BCC_1 = Chr(intBcc)
        
End Function

Private Sub Msg_Analysis(ByVal strMsg As String)
    
    Dim enqExcept   As String
    Dim strMsgType  As String * 1
    Dim lngEnq_Pos  As Long
    Dim msgTemp()   As String
    Dim strSid()    As String
    Dim msgType     As String
    Dim i           As Long
    
    Static SID         As String
    
    If strMsg = "" Then Exit Sub
    'ENQ 위치를 찾아 ENQ 부터 문자열을 취함
    lngEnq_Pos = InStr(1, strMsg, MSG_ENQ)
    If lngEnq_Pos > 0 Then
        enqExcept = Mid(strMsg, lngEnq_Pos)
    Else
        enqExcept = strMsg
    End If
    
    msgTemp = Split(enqExcept, MSG_CRLF)
    
    For i = LBound(msgTemp) To UBound(msgTemp)
        strMsgType = Mid(Replace(msgTemp(i), MSG_ENQ, ""), 3, 1)
        Debug.Print msgTemp(i)
        Select Case strMsgType
            Case MSG_H      ' "H" 'Header
            Case MSG_P      ' "P" 'Patient Informaiton
            Case MSG_Q      ' "Q" 'Request Information
                Set Order = New clsMsg_Query
                With Order
                    .MSG_QUERY = msgTemp(i)
                End With
            Case MSG_O      ' "O" 'Test Order
                strSid = Split(msgTemp(i), DLM_F)
                If UBound(strSid) > 2 Then
                    SID = strSid(2)
                End If
                If (Trim(SID) = "") And UBound(strSid) > 3 Then
                    SID = Replace(strSid(3), DLM_C, "-")
                End If
                
                If UBound(strSid) > 11 Then
                    Select Case Trim(strSid(11))
                        Case ""
                            msgType = MSG_GEN
                        Case MSG_QCT
                            msgType = MSG_QCT
                        Case Else
                            msgType = MSG_ETC
                    End Select
                Else
                    msgType = MSG_GEN
                End If
            
            Case MSG_R      ' "R" 'Result
                Set Result = New clsMsg_Result
                With Result
                    .MSG_SID = SID
                    .MSG_TYPE = msgType
                    .MSG_RESULT = msgTemp(i)
                End With
                Set Result = Nothing
            Case MSG_C      ' "C" 'Comment
            Case MSG_M      ' "M" 'Manufacturer Information
            Case MSG_S      ' "S" 'Scientific
            Case MSG_L      ' "L" 'Message Terminator
            Case Else
            
        End Select
    Next
End Sub

Private Sub Command1_Click()
Call comEQP_OnComm
End Sub

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    'txtDate.Text = Format(Now, "yyyy-mm-dd")
    txtWorkSpcno.Text = ""
    
    Call cmdClear_Click         ' 초기화
    Call SetListHeader          ' 리스트해더
    
    Call Get_Setting            ' 통신설정
    Call SetItem_List           ' 검사항목
    Call Set_ComCharacter       ' 통신문자
    
    Call cmdRun_Click           ' 실행
End Sub

Private Sub Get_Setting()
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub Get_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True


'    With comEQP
'        .CommPort = 1
'        '속도,페리티,테이타비트,stop bit
'        .Settings = "9600,n,8,1"
'        .PortOpen = True
'        .RTSEnable = True
'        .RThreshold = 1
'        .SThreshold = 1
'    End With


            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")

            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
                '.Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
                '.InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
                '.DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
                '.EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
                '.NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                '.RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
                '.InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
                '.RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
                '.SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
        
        .RTSEnable = True
        .RThreshold = 1
        .SThreshold = 1
            
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdStop_Click
    Set Result = Nothing
End Sub

Private Sub imgPort_DblClick()
    If lvwCuData.Visible Then
        lvwCuData.Visible = False
    Else
        lvwCuData.Visible = True
    End If
    
End Sub

Private Sub imgReceive_DblClick()
    If pnlCom2.Visible = True Then
        pnlCom2.Visible = False
    Else
        pnlCom2.Visible = True
    End If
End Sub

Private Sub imgSend_DblClick()
    If pnlCom.Visible = True Then
        pnlCom.Visible = False
    Else
        pnlCom.Visible = True
    End If
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub

Private Sub lvwError_Click()
    Dim itemX As ListItem
    Set itemX = lvwError.SelectedItem
    If Not itemX Is Nothing Then
        txtEqpno = itemX.Text
        txtSpcno.SetFocus
    End If

End Sub

Private Sub lvwWorkList_DblClick()
    
'    If MsgBox("선택한 검체번호를 삭제할까요?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
'        lvwWorkList.ListItems.Remove 3
'    End If

End Sub


Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub Result_MsgBegin(ByVal SID As String)
    
    Dim itemX As ListItem
    
    Set itemX = lvwComplete.FindItem(Trim(SID), lvwTag, , lvwWhole)
    If itemX Is Nothing Then
        Set itemX = lvwComplete.ListItems.Add(, , Trim(SID))
        If Not itemX Is Nothing Then
            With itemX
                .Key = COL_KEY & Trim(SID)
                .Tag = Trim(SID)
                .SmallIcon = "ITM"
            End With
        End If
    End If
        
End Sub

Private Sub Result_Update(ByVal Result As clsResult)
    Dim objWorkList    As clsCommon
    Dim objUpdate   As clsEqpResult
    Dim itemX       As ListItem
    Dim itemA       As ListItem
    Dim itemS       As ListSubItem
    CallForm = "frmComm - Private Sub Result_Update(ByVal Result As clsResult)"
On Error GoTo ErrorRoutine

    Set objUpdate = New clsEqpResult
    '검사 결과 업데이트
    With objUpdate
        Set mAdoRs = New ADODB.Recordset
        
        Set objWorkList = New clsCommon
        Call objWorkList.SetAdoCn(AdoCn_SQL)
        Set mAdoRs = objWorkList.Get_WorkList(Trim(Result.Rst_Sid))
        
        'If .Spc_Exists(Result.Rst_Sid) Then
        If Not mAdoRs Is Nothing Then
            .SPCID = Trim(Result.Rst_Sid)   '검체번호
            .EQPNUM = Result.Rst_Eid        '장비번호
            .SPCTYPE = Result.Rst_Type      '메시지 타입
            'Call .Set_EqpResultjet(Result.Rst_Tag, Result.Rst_Values)   '로칼디비에 등록
            Call objUpdate.SetAdoCn(AdoCn_SQL)
            Call .Set_EqpResultsql(Result.Rst_Sid, Mid(Result.Rst_Test, 2), Result.Rst_Values, Result.Rst_Eid, Result.Rst_Error)   '서버에 등록
            
            'Call .Set_EqpResultsql(Result.Rst_Tag, Result.Rst_Values, Result.Rst_Error) '서버에 등록
            
            '결과 완료테이블에 추가
            Set itemX = lvwComplete.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
            If itemX Is Nothing Then
                Set itemX = lvwComplete.ListItems.Add()
                With itemX
                    .Key = COL_KEY & Result.Rst_Sid '아이템 키에 검체번호
                    .Text = Result.Rst_Sid          '아이템 에 검체번호
                    .Tag = Result.Rst_Type          '테그에 결과 타입
                    .SmallIcon = "ITM"
                End With
            End If
            '결과값 등록 fCellItem()
            itemX.SubItems(lvwComplete.ColumnHeaders(Result.Rst_Test).SubItemIndex) = Result.Rst_Values
            Set itemS = itemX.ListSubItems(lvwComplete.ColumnHeaders(Result.Rst_Test).SubItemIndex)
            itemS.Tag = Result.Rst_Error '서브아이템 테그에 에러 메시지
            Set itemS = Nothing
            Set itemX = Nothing
        Else '검체번호가 없는것은 미등록 테이블에
            Set itemX = lvwError.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
            If itemX Is Nothing Then
                Set itemX = lvwError.ListItems.Add()
                With itemX
                    .Key = COL_KEY & Result.Rst_Sid
                    .Text = Result.Rst_Sid
                    .Tag = Result.Rst_Type
                    .SmallIcon = "LSE"
                End With
            End If
            itemX.SubItems(lvwError.ColumnHeaders(Result.Rst_Test).SubItemIndex) = Result.Rst_Values
            Set itemS = itemX.ListSubItems(lvwError.ColumnHeaders(Result.Rst_Test).SubItemIndex)
            itemS.Tag = Result.Rst_Error
            Set itemS = Nothing
            Set itemX = Nothing
        End If
    End With
    
    Set objUpdate = Nothing
Exit Sub
ErrorRoutine:
    Set objUpdate = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Result_MsgSplit(ByVal Result As clsResult)

    Dim itemX As ListItem
    Dim itemH As ColumnHeader
    
    '메치 테이블에서 검사코드를 가져옴
    Set itemX = lvwCuData.FindItem(Trim(Result.Rst_Eid), lvwTag, , lvwWhole)
    If Not itemX Is Nothing Then
        'Result.Rst_Tag = Trim(itemX.SubItems(1))
        Result.Rst_Tag = Trim(itemX.SubItems(2))
        
        Call Result_Update(Result)
        Set itemX = Nothing
    End If
    '검사코드가 없는것은 등록 하지 않음
    
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

' 통신상태 확인 관련이벤트
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.Text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim i               As Long
    Dim lngFIleNum      As Long
    Dim strTemp         As String
    Dim strTemp2        As String
    Dim bteBuffer()     As Byte
    
On Error GoTo ErrorRoutine
    
    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowOpen
        lngFIleNum = FreeFile
        
        Open .FileName For Binary Access Read As #lngFIleNum
        
        txtCOM2.Text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.Text = strTemp
                
        Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum
        
End Sub

Private Sub cmdCOMSave_Click()
    Dim lngFIleNum      As Long
    
On Error GoTo ErrorRoutine

    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowSave
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum
        Print #lngFIleNum, _
              Format(Date, "YYYY년 MM월 DD일") & "  "; Time & vbNewLine & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine & _
              txtCom.Text & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine
    Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum

End Sub

Private Sub cmdCOMOutput_Click()
    'Call COM_OUTPUT(StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode))
    Call COM_OUTPUT(charCOM_Convert(txtCom.SelText))
End Sub

Private Sub cmdCOMClear_Click()
    mlngRecLen = 0
    txtCom.Text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.Text = ""
End Sub

Private Sub cmdCOMInput_Click()
    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

    Call ComReceive(bytTemp)
End Sub

Private Sub cmdCOMInput2_Click()
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.Text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(bytTemp)
End Sub

Private Sub cmdCOMOutput2_Click()
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.Text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트



Private Sub txtWorkSpcno_KeyPress(KeyAscii As Integer)
    Dim itemH           As ColumnHeader
    Dim objWorkList     As clsCommon
    Dim itemX           As ListItem
    Dim itmFound        As ListItem
    
    If KeyAscii = vbKeyReturn Then
        Set objWorkList = New clsCommon
        With objWorkList
            .SetAdoCn AdoCn_SQL
            Set mAdoRs = .Get_WorkList(Trim(txtWorkSpcno.Text))
        End With
        
        Set objWorkList = Nothing
        
        If Not mAdoRs Is Nothing Then
            If Not mAdoRs.BOF Then
                Do Until mAdoRs.EOF
                    Set itmFound = lvwWorkList.FindItem(Trim(mAdoRs.Fields("BARCODENUMBER")), 2, , 1)
                    If itmFound Is Nothing Then
                        Set itemX = lvwWorkList.ListItems.Add(, , Trim(mAdoRs.Fields("BARCODENUMBER")))
                        With itemX
                            .Key = COL_KEY & Trim(mAdoRs.Fields("BARCODENUMBER"))
                            .Tag = Trim(mAdoRs.Fields("BARCODENUMBER"))
                            .SmallIcon = "ITM"
                        End With
                    Else
                        'MsgBox "이미 등록된 검체 번호입니다", vbInformation, Me.Caption
                    End If
                    mAdoRs.MoveNext
                Loop
            Else
                MsgBox "접수되지 않은 검체 번호입니다", vbInformation, Me.Caption
            End If
        End If
        txtWorkSpcno.Text = ""
        txtWorkSpcno.SetFocus
        Set mAdoRs = Nothing
    End If

End Sub

'Private Sub WorkList_MsgBegin(ByVal strData As String)
'
'    Dim strSid  As String
'
'    Dim itemX As ListItem
'
'    Set itemX = lvwWorkList.FindItem(Trim(SID), lvwTag, , lvwWhole)
'    If itemX Is Nothing Then
'        Set itemX = lvwWorkList.ListItems.Add(, , Trim(SID))
'        If Not itemX Is Nothing Then
'            With itemX
'                .Key = COL_KEY & Trim(SID)
'
'                .Tag = Trim(SID)
'                .SmallIcon = "ITM"
'            End With
'        End If
'    End If
'
'
'End Sub


Private Sub WorkList_MsgSplit(ByVal Result As clsResult)

    Dim itemX As ListItem
    Dim itemH As ColumnHeader
    
    '메치 테이블에서 검사코드를 가져옴
    Set itemX = lvwWorkList.FindItem(Trim(Result.Rst_Test), lvwTag, , lvwWhole)
    If Not itemX Is Nothing Then
        Result.Rst_Tag = Trim(itemX.SubItems(1))
'        Call Result_Update(Result)
        Set itemX = Nothing
    End If
    '검사코드가 없는것은 등록 하지 않음

End Sub


