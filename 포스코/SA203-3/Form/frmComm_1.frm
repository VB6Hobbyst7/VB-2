VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmComm_1 
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
      TabIndex        =   11
      Top             =   570
      Width           =   11925
      Begin TabDlg.SSTab tabWork 
         Height          =   5850
         Left            =   45
         TabIndex        =   12
         Top             =   0
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   10319
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   " 받은 결과 "
         TabPicture(0)   =   "frmComm_1.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdQuery"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "pnlCom2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "pnlCom"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdServer"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lvwComplete"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "mskYear"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtData"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lvwData"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lvwCuData"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   " 저장 결과 "
         TabPicture(1)   =   "frmComm_1.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvwError"
         Tab(1).Control(1)=   "mskDate"
         Tab(1).Control(2)=   "cmdAppend"
         Tab(1).Control(3)=   "Label10"
         Tab(1).ControlCount=   4
         Begin MSComctlLib.ListView lvwCuData 
            Height          =   5415
            Left            =   2250
            TabIndex        =   39
            Top             =   450
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
         Begin MSComctlLib.ListView lvwData 
            Height          =   4965
            Left            =   7245
            TabIndex        =   40
            Top             =   810
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   8758
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
         Begin VB.TextBox txtData 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   5805
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   6000
         End
         Begin MSMask.MaskEdBox mskYear 
            Height          =   285
            Left            =   1350
            TabIndex        =   34
            Top             =   465
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ListView lvwError 
            Height          =   4995
            Left            =   -74955
            TabIndex        =   14
            Top             =   765
            Width           =   11715
            _ExtentX        =   20664
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
         Begin MSComctlLib.ListView lvwComplete 
            Height          =   4980
            Left            =   90
            TabIndex        =   13
            Top             =   810
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   8784
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin HSCotrol.CButton cmdServer 
            Height          =   330
            Left            =   10665
            TabIndex        =   35
            Top             =   450
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            Caption         =   "서버등록"
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
         Begin MSMask.MaskEdBox mskDate 
            Height          =   285
            Left            =   -73785
            TabIndex        =   37
            Top             =   470
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin HSCotrol.CButton cmdAppend 
            Height          =   330
            Left            =   -64335
            TabIndex        =   38
            Top             =   405
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            Caption         =   "서버등록"
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
         Begin HSCotrol.UserPanel pnlCom 
            Height          =   5400
            Left            =   -675
            TabIndex        =   24
            Top             =   2160
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
               TabIndex        =   26
               Top             =   4650
               Width           =   11610
               Begin HSCotrol.CButton cmdCOMSave 
                  Height          =   360
                  Left            =   10515
                  TabIndex        =   27
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
                  TabIndex        =   28
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
                  TabIndex        =   29
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
                  TabIndex        =   30
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
               TabIndex        =   25
               Top             =   270
               Width           =   11595
            End
         End
         Begin HSCotrol.UserPanel pnlCom2 
            Height          =   5385
            Left            =   2880
            TabIndex        =   16
            Top             =   2385
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
            Begin VB.Frame Frame2 
               Height          =   645
               Left            =   60
               TabIndex        =   18
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
                  TabIndex        =   19
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
                  TabIndex        =   20
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
                  TabIndex        =   21
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
                  TabIndex        =   22
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
                  Left            =   4635
                  TabIndex        =   23
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
                  TabIndex        =   31
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
                  TabIndex        =   32
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
               TabIndex        =   17
               Top             =   405
               Width           =   5730
            End
         End
         Begin HSCotrol.CButton cmdQuery 
            Height          =   330
            Left            =   9585
            TabIndex        =   41
            Top             =   450
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            Caption         =   "조 회"
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
            Caption         =   "검체접수년 :"
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
            Left            =   135
            TabIndex        =   33
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "검사자료 : "
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
            Left            =   -74865
            TabIndex        =   15
            Top             =   540
            Width           =   1005
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
            Picture         =   "frmComm_1.frx":0038
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":05D2
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":0B6C
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":1106
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":16A0
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":1C3A
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
            Picture         =   "frmComm_1.frx":21D4
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":276E
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":2D08
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":32A2
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3B34
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3C8E
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_1.frx":3DE8
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
      TabIndex        =   1
      Top             =   6495
      Width           =   11940
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6375
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   10
         Top             =   225
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
         TabIndex        =   9
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm_1.frx":3F42
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
         TabIndex        =   8
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   9105
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   8040
         TabIndex        =   6
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11010
         Picture         =   "frmComm_1.frx":51C4
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_1.frx":574E
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_1.frx":5CD8
         Top             =   255
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmComm_1"
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

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult
Attribute Result1.VB_VarHelpID = -1

Private mAdoRs      As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer As String
Private f_strSend   As String

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Sub f_subGet_받은결과표시()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem
    
    CallForm = "frmComm - Private Sub cmdQuery_Click()"
    
On Error GoTo ErrorRoutine
    Me.MousePointer = 11
    
    '-- CLEAR
    Do
        Set itemX = lvwError.SelectedItem
        If itemX Is Nothing Then Exit Do
        lvwError.ListItems.Remove (itemX.Index)
    Loop
    Set itemX = Nothing
    
    sqlDoc = "select SPCNO, TESTCD, RSTVAL" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskDate.Text & "'" & _
             " order by TRANSTM"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
    
        Set itemX = lvwError.FindItem(Trim$(adoRS(0) & ""), lvwText, , lvwWhole)
        If itemX Is Nothing Then
            Set itemX = lvwError.ListItems.Add(, , Trim$(adoRS(0) & ""))
            If Not itemX Is Nothing Then
                With itemX
                    .Key = COL_KEY & Trim$(adoRS(0) & "")
                    .Tag = "G"
                    .Text = Trim$(adoRS(0) & "")
                    .SmallIcon = "LSE"
                End With
            End If
        End If
        
'        If Mid$(adoRS(0) & "", 9, 2) = "PC" And Trim$(adoRS(1) & "") = "06A" Then
'            itemX.SubItems(lvwError.ColumnHeaders(COL_KEY & "XXX").SubItemIndex) = Trim$(adoRS(2) & "")
'        Else
            itemX.SubItems(lvwError.ColumnHeaders(COL_KEY & Trim$(adoRS(1) & "")).SubItemIndex) = Trim$(adoRS(2) & "")
'        End If
        
        Set itemX = Nothing
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
    Me.MousePointer = 0
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)
    Me.MousePointer = 0

End Sub

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef strSpec As String, ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "   "
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ,0": strPcFlag = "PC "
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ,0,"
            End If
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub

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
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC_L", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC_H", (lvwCuData.Width - 310) * 0.2)
        End With
        .HideColumnHeaders = False
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
    
    '-- 과거데이타
    With lvwData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, "TEST_EQP", "장비코드", (lvwCuData.Width - 310) * 0)
            Call .Add(, "TEST_NM", "검사명", (lvwCuData.Width - 310) * 0.35)
            Call .Add(, "TSET_OLD_LIST", "직전결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "TEST_CUR_LIST", "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DETAL", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "PANIC", "PANIC", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
End Sub

Private Sub SetlistView_Complete(lvw As Listview)
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Set objHeadeItem = New clsCommon
    With objHeadeItem
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_TestItemList(INS_CODE)
    End With
    Set objHeadeItem = Nothing
    
    If Not mAdoRs Is Nothing Then
        With lvw
            .ColumnHeaders.Clear
            Call .ColumnHeaders.Add(, "EQP_ID", "검체 번호")
            Do Until mAdoRs.EOF
                Set itemH = .ColumnHeaders.Add
                With itemH
                    '컬럽 헤더키를 장비검사 코드로
                    .Key = COL_KEY & Trim(mAdoRs.Fields("TESTCD_EQP") & "")
                    '컬럽명은 검사 항목 이름
                    .Text = Trim(mAdoRs.Fields("TESTNM") & "")
                    '테그는 검사 코드로
                    .Tag = Trim(mAdoRs.Fields("TESTCD") & "")
                End With
                Set itemH = Nothing
                mAdoRs.MoveNext
            Loop
        End With
    End If
    Set mAdoRs = Nothing
End Sub

Private Sub SetItem_List()

    Dim objItem As clsCommon
    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub SetItem_List()"
    Set objItem = New clsCommon
    '검사 코드 메치테이블
    With objItem
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_TestItemList(INS_CODE)
    End With
    
    lvwCuData.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        Do Until mAdoRs.EOF
            Set itemX = lvwCuData.ListItems.Add(, , Trim(mAdoRs.Fields("TESTCD_EQP") & ""), , "LST")
                itemX.SubItems(1) = Trim(mAdoRs.Fields("TESTCD") & "")
                itemX.SubItems(2) = Trim(mAdoRs.Fields("TESTNM") & "")
                itemX.SubItems(3) = ""
                itemX.SubItems(4) = Trim(mAdoRs.Fields("DELTA") & "")
                itemX.SubItems(5) = Trim(mAdoRs.Fields("DELTAGBN") & "")
                itemX.SubItems(6) = Trim(mAdoRs.Fields("PANICL") & "")
                itemX.SubItems(7) = Trim(mAdoRs.Fields("PANICH") & "")
                itemX.Tag = Trim(mAdoRs.Fields("TESTCD_EQP") & "")
            Set itemX = Nothing
            
            Set itemA = lvwData.ListItems.Add(, , Trim(mAdoRs.Fields("TESTCD_EQP") & ""), , "LST")
                itemA.SubItems(1) = Trim(mAdoRs.Fields("TESTNM") & "")
                itemA.Tag = Trim(mAdoRs.Fields("TESTCD_EQP") & "")
            Set itemA = Nothing
            
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

Private Sub cmdAppend_Click()
    
    Dim itemLX  As ListItem:    Dim itemSX  As ListSubItem
    Dim itemLA  As ListItem
    Dim objSave As clsEqpResult
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strChatno   As String, strOrdcd As String
    
    CallForm = "frmComm - Private Sub cmdServer_Click()"
On Error GoTo ErrorRoutine
    
    For Each itemLX In lvwError.ListItems
        If itemLX.SmallIcon = "ITM" Then
            Set objSave = New clsEqpResult
            With objSave
                .EQPNUM = itemLX.Text
                .SPCID = itemLX.Text
                
                .SPCTYPE = itemLX.Tag
                For Each itemSX In itemLX.ListSubItems
                    '서브아이템에 검사 결과 가 있으면
                    If Trim(itemSX.Text) <> "" Then
    '                    Call .Set_EqpResultjet(lvwComplete.ColumnHeaders(itemSX.Index + 1).Tag, itemSX.Text)
                        
                        strChatno = "": strOrdcd = ""
                        strOrdcd = lvwComplete.ColumnHeaders(itemSX.Index + 1).Tag
                        
                        sqlDoc = "select distinct CHART_NO, ORD_CODE from L3A01" & _
                                 " where SAMPLE_DATE = '" & Mid$(mskYear.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'" & _
                                 "   and SAMPLE_SEQ  = '" & Mid$(itemLX.Text, 5, 3) & "'" & _
                                 "   and PART        = '" & Mid$(itemLX.Text, 8, 2) & "'" & _
                                 "   and ORD_CODE   in ( '" & strOrdcd & "')"
                        adoRS.CursorLocation = adUseClient
                        adoRS.Open sqlDoc, AdoCn_SQL
                        If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                        If Not adoRS.EOF Then strChatno = Trim$(adoRS(0) & "") ': strOrdcd = Trim$(adoRS(1) & "")
                        adoRS.Close:    Set adoRS = Nothing
                        
                        sqlDoc = "exec p_l3a01interface" & _
                                "      'U'," & _
                                "      '" & Mid$(mskYear.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'," & _
                                "      '" & Mid$(itemLX.Text, 8, 2) & "'," & _
                                "      '" & strChatno & "'," & _
                                "      '" & Mid$(itemLX.Text, 5, 3) & "'," & _
                                "      '" & strOrdcd & "'," & _
                                "      '0', '" & itemSX.Text & "'"
                        AdoCn_SQL.Execute sqlDoc
                    End If
                Next
                
                Set itemLA = lvwError.FindItem(itemLX.Text, lvwText, , lvwWhole)
                If Not itemLA Is Nothing Then itemLA.SmallIcon = "LSE"
                Set itemLA = Nothing
            End With
        
            Set itemSX = Nothing
            Set objSave = Nothing
        End If
    Next
    Set itemLX = Nothing

Exit Sub
ErrorRoutine:
    Set itemLX = Nothing
    Set itemLA = Nothing
    Set itemSX = Nothing
    Set objSave = Nothing
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
    Dim itemS As ListSubItem
    
    For Each itemX In lvwCuData.ListItems
        itemX.SubItems(3) = ""
    Next
    Set itemX = Nothing
    
    Do
        Set itemX = lvwComplete.SelectedItem
        If itemX Is Nothing Then Exit Do
        lvwComplete.ListItems.Remove (itemX.Index)
    Loop
    Set itemX = Nothing
    
    Do
        Set itemX = lvwError.SelectedItem
        If itemX Is Nothing Then Exit Do
        lvwError.ListItems.Remove (itemX.Index)
    Loop
    Set itemX = Nothing

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
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
        
    End If
        
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

Private Sub cmdQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem
    
    CallForm = "frmComm - Private Sub cmdQuery_Click()"
    
On Error GoTo ErrorRoutine
    Me.MousePointer = 11
    
    '-- CLEAR
    Do
        Set itemX = lvwComplete.SelectedItem
        If itemX Is Nothing Then Exit Do
        lvwComplete.ListItems.Remove (itemX.Index)
    Loop
    Set itemX = Nothing
    
    sqlDoc = "select SPCNO, TESTCD, RSTVAL" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskDate.Text & "'" & _
             " order by TRANSTM"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
    
        Set itemX = lvwComplete.FindItem(Trim$(adoRS(0) & ""), lvwText, , lvwWhole)
        If itemX Is Nothing Then
            Set itemX = lvwComplete.ListItems.Add(, , Trim$(adoRS(0) & ""))
            If Not itemX Is Nothing Then
                With itemX
                    .Key = COL_KEY & Trim$(adoRS(0) & "")
                    .Tag = "G"
                    .Text = Trim$(adoRS(0) & "")
                    .SmallIcon = "LSE"
                End With
            End If
        End If
        
'        If Mid$(adoRS(0) & "", 9, 2) = "PC" And Trim$(adoRS(1) & "") = "06A" Then
'            itemX.SubItems(lvwComplete.ColumnHeaders(COL_KEY & "XXX").SubItemIndex) = Trim$(adoRS(2) & "")
'        Else
            itemX.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Trim$(adoRS(1) & "")).SubItemIndex) = Trim$(adoRS(2) & "")
'        End If
        
        Set itemX = Nothing
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
    Me.MousePointer = 0
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)
    Me.MousePointer = 0
    
End Sub

Private Sub cmdServer_Click()
    
    Dim itemLX  As ListItem:    Dim itemSX  As ListSubItem
    Dim itemLA  As ListItem:    Dim itemSA  As ListSubItem
    Dim objSave As clsEqpResult
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strChatno   As String, strOrdcd As String
    
    CallForm = "frmComm - Private Sub cmdServer_Click()"
On Error GoTo ErrorRoutine
    
    Do
        Set itemLX = lvwComplete.SelectedItem
        If itemLX Is Nothing Then Exit Do
        
        Set objSave = New clsEqpResult
        With objSave
            .EQPNUM = itemLX.Text
            .SPCID = itemLX.Text
            
            .SPCTYPE = itemLX.Tag
            For Each itemSX In itemLX.ListSubItems
                '서브아이템에 검사 결과 가 있으면
                If Trim(itemSX.Text) <> "" Then
'                    Call .Set_EqpResultjet(lvwComplete.ColumnHeaders(itemSX.Index + 1).Tag, itemSX.Text)
                    
                    strChatno = "": strOrdcd = ""
'                    If lvwComplete.ColumnHeaders(itemSX.Index + 1).Key = COL_KEY + "06A" Then
'                        strOrdcd = lvwComplete.ColumnHeaders(itemSX.Index + 1).Tag & "','" & _
'                                   lvwComplete.ColumnHeaders(lvwComplete.ColumnHeaders(COL_KEY & "XXX").SubItemIndex + 1).Tag & ""
'                    Else
                        strOrdcd = lvwComplete.ColumnHeaders(itemSX.Index + 1).Tag
'                    End If
                    sqlDoc = "select distinct CHART_NO, ORD_CODE from L3A01" & _
                             " where SAMPLE_DATE = '" & Mid$(mskYear.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'" & _
                             "   and SAMPLE_SEQ  =  " & Format(Mid$(itemLX.Text, 5, 3), "##0") & "" & _
                             "   and PART        = '" & Mid$(itemLX.Text, 8, 2) & "'" & _
                             "   and ORD_CODE   in ( '" & strOrdcd & "')"
                    adoRS.CursorLocation = adUseClient
                    adoRS.Open sqlDoc, AdoCn_SQL
                    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                    If Not adoRS.EOF Then strChatno = Trim$(adoRS(0) & "") ': strOrdcd = Trim$(adoRS(1) & "")
                    adoRS.Close:    Set adoRS = Nothing
                    
                    sqlDoc = "exec p_l3a01interface" & _
                            "      'U'," & _
                            "      '" & Mid$(mskYear.Text, 1, 4) + Mid$(itemLX.Text, 1, 4) & "'," & _
                            "      '" & Mid$(itemLX.Text, 8, 2) & "'," & _
                            "      '" & strChatno & "'," & _
                            "       " & Format(Mid$(itemLX.Text, 5, 3), "##0") & "," & _
                            "      '" & strOrdcd & "'," & _
                            "      '0', '" & itemSX.Text & "'"
                    AdoCn_SQL.Execute sqlDoc
                End If
            Next
            
            '검사결과 등록 테이블에 등록
            Set itemLA = lvwError.FindItem(Trim(itemLX.Text), lvwText, , lvwWhole)
            If itemLA Is Nothing Then
                Set itemLA = lvwError.ListItems.Add()
            End If
            With itemLA
                .Key = COL_KEY & Trim(itemLX.Text)
                .Text = Trim(itemLX.Text)
                .Tag = itemLX.Tag
                .SmallIcon = "LSE"
                For Each itemSA In itemLX.ListSubItems
                    .SubItems(itemSA.Index) = itemSA.Text
                Next
                Set itemSA = Nothing
            End With
            Set itemLA = Nothing
    
            '미등록 테이블에서 제거
            Call lvwComplete.ListItems.Remove(itemLX.Index)
            
        End With
    
        Set itemSX = Nothing
        Set objSave = Nothing
    Loop
    Set itemLX = Nothing

Exit Sub
ErrorRoutine:
    Set itemLX = Nothing
    Set itemLA = Nothing
    Set itemSX = Nothing
    Set objSave = Nothing
    Call ErrMsgProc(CallForm)

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

Private Sub ComReceive(ByRef RecData() As Byte)
    
    Dim strTemp     As String
    Dim strRec      As String
    Dim intEOT_POS  As Long
    
    Dim intIdx  As Integer
    Dim strOrder As String, strSpec  As String, strPcFlag   As String
    
    Static OrgMsg As String
    strRec = StrConv(RecData, vbUnicode)
    
    Print #1, strRec;
    
    Call COM_INPUT(strRec)
    
    txtData.Text = txtData.Text + strRec
    
    For intIdx = 1 To Len(strRec)
        Select Case Mid$(strRec, intIdx, 1)
            Case Chr(1)     '-- SOH
                            Call COM_OUTPUT(Chr(6))
                            
            Case Chr(3)     '-- ETX
                            Call COM_OUTPUT(Chr(4))
                        
            Case Chr(4)     '-- EOT
                            If f_strSend <> "" Then Call COM_OUTPUT(Chr(1))
                        
            Case Chr(5)     '-- ENQ
                            Call COM_OUTPUT(Chr(21))
            Case Chr(6)     '-- ACK
                            Call COM_OUTPUT(f_strSend)
                        
            Case Chr(10)    '-- LF
                            f_strBuffer = f_strBuffer + Mid$(strRec, intIdx, 1)
                            Select Case Mid$(f_strBuffer, 5, 6)
                                Case "701,06"   '-- JOB LIST 요청
                                                Call f_subGet_JobList(Mid$(mskYear.Text, 1, 4) + Mid$(f_strBuffer, 12, 11), strOrder, strSpec, strPcFlag)
                                                f_strSend = Mid$(f_strBuffer, 1, 7) + ",01, 0, 0,1,ST," & strSpec & "," + _
                                                           Mid$(f_strBuffer, 12, 9) + strPcFlag + "," + Space(20) + "," + Space(25) + "," + Space(25) + "," + _
                                                           Space(18) + "," + Space(15) + ", ," + Space(12) + "," + Space(18) + "," + _
                                                           Format(Now, "ddmmyy") + "," + Format(Now, "hhmm") + "," + Space(20) + ",000,5," + _
                                                           Space(6) + ",F," + Space(25) + "," + Space(7) + "," + Space(4) + "," + Space(4) + _
                                                           Space(6) + "," + "001," + strOrder + "]"
                                                f_strSend = f_strSend + f_funGet_CheckSum(f_strSend) + Chr(13) + Chr(10)
                                                
                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                                                
                                Case "701,02"   '-- job list 받은 후
                                                If Trim$(Mid$(f_strBuffer, 12, 2)) = "0" Then f_strSend = ""
                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                                                
                                Case "702,01"   '-- 검사결과(cup header)
                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                                                
                                Case "702,03"   '-- 검사결과
                                                Set Result1 = New clsResult
                                                With Result1
                                                    .Rst_Sid = Trim$(Mid$(f_strBuffer, 48, 11))
                                                    .Rst_Eid = Mid$(f_strBuffer, 102, 1)
                                                    .Rst_Type = "G"
                                                    .Rst_Test = Mid$(f_strBuffer, 60, 4)
                                                    .Rst_Values = Mid$(f_strBuffer, 82, 9)
                                                    .Rst_Tag = Mid$(f_strBuffer, 60, 4)
                                                    .Rst_Error = ""
                                                End With
'                                                Call Result_MsgBegin(Trim$(Mid$(f_strBuffer, 48, 9)))
                                                Call Result_MsgSplit(Result1)
                                                Set Result1 = Nothing

                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                                                
                                Case "702,05"   '-- end of cup
                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                                                
                                Case Else
                                                Call COM_OUTPUT(Chr(3))
                                                f_strBuffer = ""
                            End Select
                            txtData.Text = ""
                            
            Case Else:      f_strBuffer = f_strBuffer + Mid$(strRec, intIdx, 1)
                
        End Select
    Next

End Sub




Public Function f_funGet_CheckSum(ByVal strPara As String) As String

    Dim intIdx      As Integer
    Dim intChkSum   As Integer
    
    intChkSum = 0
    For intIdx = 1 To Len(strPara)
        intChkSum = intChkSum + Asc(Mid$(strPara, intIdx, 1))
    Next
    
    intChkSum = 256 - intChkSum Mod 256
    
    f_funGet_CheckSum = Hex(intChkSum)
        
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

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    Call cmdClear_Click         ' 초기화
    Call SetListHeader          ' 리스트해더
    Call Get_Setting            ' 통신설정
    Call SetItem_List           ' 검사항목
    Call Set_ComCharacter       ' 통신문자
    
    Call cmdRun_Click           ' 실행
    
    mskYear.Text = Format$(Now, "YYYYMMDD")
    mskDate.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "dump_job.log" For Append As #1

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
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
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
    
    Close #1

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

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        pnlCom.Visible = False
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        pnlCom.Visible = True:  pnlCom.ZOrder 0
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub

Private Sub lvwComplete_Click()

    Dim itemX   As ListItem
    
    Set itemX = lvwComplete.SelectedItem
    If Not itemX Is Nothing Then
        If itemX.SmallIcon = "LSE" Then
            itemX.SmallIcon = "ITM"
        Else
            itemX.SmallIcon = "LSE"
        End If
    End If
    Set itemX = Nothing
    
End Sub


Private Sub lvwComplete_DblClick()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem:    Dim itemA   As ListItem:    Dim itemL   As ListItem
    Dim itemS   As ListSubItem: Dim itemSA  As ListSubItem
    
    Dim strSample_dt    As String, strSample_no As String, strPart      As String
    Dim strOrd_cd       As String, strChart_no  As String, strResult    As String
    Dim strValue1       As String, strValue2    As String
    
    For Each itemX In lvwData.ListItems
        itemX.SubItems(2) = ""
        itemX.SubItems(3) = ""
        itemX.SubItems(4) = ""
        itemX.SubItems(5) = ""
    Next
    Set itemX = Nothing

    Set itemX = lvwComplete.SelectedItem
    If Not itemX Is Nothing Then
        For Each itemS In itemX.ListSubItems
            If Trim(itemS.Text) <> "" Then
                strSample_dt = Mid$(mskYear.Text, 1, 4) + Mid$(itemX.Text, 1, 4)
                strSample_no = Format$(Mid$(itemX.Text, 5, 3), "##0")
                strPart = Mid$(itemX.Text, 8, 2)
                strOrd_cd = lvwComplete.ColumnHeaders(itemS.Index + 1).Tag
                
                Set itemL = lvwData.FindItem(Mid$(lvwComplete.ColumnHeaders(itemS.Index + 1).Key, 2), lvwTag, , lvwWhole)
                Set itemA = lvwCuData.FindItem(Mid$(lvwComplete.ColumnHeaders(itemS.Index + 1).Key, 2), lvwTag, , lvwWhole)
                If Not itemL Is Nothing Then
                    itemL.SubItems(3) = itemS.Text
                    sqlDoc = "select c.RESULT" & _
                             "  from (select CHART_NO FROM L3A01" & _
                             "        where  SAMPLE_DATE = '" & strSample_dt & "'" & _
                             "        and    SAMPLE_SEQ  = '" & strSample_no & "'" & _
                             "        and    PART        = '" & strPart & "'" & _
                             "        and    ORD_CODE    = '" & strOrd_cd & "'" & _
                             "        group  by CHART_NO) as a," & _
                             "       (select CHART_NO, max(SAMPLE_DATE) SAMPLE_DATE from L3A01" & _
                             "        where  SAMPLE_DATE < '" & strSample_dt & "'" & _
                             "        and    ORD_CODE    = '" & strOrd_cd & "'" & _
                             "        group  by CHART_NO) AS b," & _
                             "       (select SAMPLE_DATE, CHART_NO, RESULT from L3A01" & _
                             "        where  ORD_CODE  =  '" & strOrd_cd & "'" & _
                             "        group  by SAMPLE_DATE, CHART_NO, RESULT) c" & _
                             " Where a.CHART_NO = b.CHART_NO" & _
                             "   and b.CHART_NO = c.CHART_NO" & _
                             "   and b.SAMPLE_DATE = c.SAMPLE_DATE"
                    adoRS.CursorLocation = adUseClient
                    adoRS.Open sqlDoc, AdoCn_SQL
                    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                    If Not adoRS.EOF Then itemL.SubItems(2) = Trim$(adoRS(0) & ""): strResult = Trim$(adoRS(0) & "")
                    adoRS.Close:    Set adoRS = Nothing
                    If strResult <> "" Then
                        If itemA.SubItems(4) <> "" Then
                            '-- DELTA
                            strValue1 = Abs(Val(itemS.Text) - Val(strResult))
                            strValue2 = (strValue1 / Val(strResult)) * 100
                            Select Case itemA.SubItems(5)
                                Case "1":   If strValue1 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "2":   If strValue2 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "3":   If Val(strValue2) / 30 > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                                Case "4":   If Val(strValue1) / Val(strValue2) > Val(itemA.SubItems(4)) Then itemL.SubItems(4) = "D"
                            End Select
                            '-- PANIC
                            If Val(itemS.Text) < Val(itemA.SubItems(6)) Or Val(itemS.Text) > Val(itemA.SubItems(7)) Then itemL.SubItems(5) = "P"
                        End If
                    End If
                End If
                Set itemL = Nothing
            End If
        Next
    End If
    Set itemX = Nothing

End Sub


Private Sub lvwError_DblClick()

    Dim itemX   As ListItem
    
    Set itemX = lvwError.SelectedItem
    If Not itemX Is Nothing Then
        If itemX.SmallIcon = "LSE" Then
            itemX.SmallIcon = "ITM"
        Else
            itemX.SmallIcon = "LSE"
        End If
    End If
    Set itemX = Nothing
    
End Sub

Private Sub mskDate_GotFocus()

    With mskDate
        .Tag = .Text
        .SelStart = 0
        .SelLength = Len(.Text) + 2
    End With
    
End Sub


Private Sub mskDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdQuery_Click
    ElseIf Not KeyAscii = vbKeyBack Then
        mskDate.SelLength = 1
    End If
    
End Sub



Private Sub mskYear_GotFocus()

    With mskYear
        .SelStart = 0
        .SelLength = Len(.Text) + 2
    End With '
    
End Sub


Private Sub mskYear_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskYear.SelLength = 1
    
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
    
    Dim objUpdate   As clsEqpResult
    Dim itemX       As ListItem
    Dim itemA       As ListItem
    Dim itemS       As ListSubItem
    
    CallForm = "frmComm - Private Sub Result_Update(ByVal Result As clsResult)"
On Error GoTo ErrorRoutine

    Set objUpdate = New clsEqpResult
    '검사 결과 업데이트
    With objUpdate
        If .Spc_Exists(Result.Rst_Sid) Then
            .SPCID = Trim(Result.Rst_Sid)   '검체번호
            .EQPNUM = ""                    '장비번호
            .SPCTYPE = Result.Rst_Type      '메시지 타입
            Call .Set_EqpResultjet(Result.Rst_Tag, Result.Rst_Values)   '로칼디비에 등록
'            Call .Set_EqpResultsql(Result.Rst_Tag, Result.Rst_Values, Result.Rst_Error) '서버에 등록
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
            '결과값 등록
            itemX.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
            Set itemS = itemX.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
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
            itemX.SubItems(lvwError.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
            
            Set itemS = itemX.ListSubItems(lvwError.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
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

On Error GoTo ErrorRoutine
    
    Dim sqlDoc  As String
    
    Dim itemX As ListItem
    Dim itemH As ListItem
    Dim itemS As ListSubItem
    
    CallForm = "frmComm - Private Sub Result_MsgSplit()"

    '메치 테이블에서 검사코드를 가져옴
    Set itemX = lvwCuData.FindItem(Trim(Result.Rst_Test), lvwTag, , lvwWhole)
    If Not itemX Is Nothing Then
        If Mid$(Result.Rst_Sid, 10, 2) = "PC" And Trim(Result.Rst_Test) = "06A" Then
            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
            Result.Rst_Test = "XXX"
            Result.Rst_Tag = ""
        Else
            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
            Result.Rst_Tag = Trim(itemX.SubItems(1))
        End If
        
'        Call Result_Update(Result)
        
        sqlDoc = "insert into INTERFACE003(" & _
                 "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, EQUIPCD)" & _
                 "    values( '" & Result.Rst_Sid & "', '" & Result.Rst_Test & "'," & _
                 "            '" & Result.Rst_Eid & "', '" & Format$(Now, "YYYYMMDD") & "'," & _
                 "            '" & Format$(Now, "MMSS") & "', '" & Result.Rst_Values & "'," & _
                 "            '" & INS_CODE & "')"
        AdoCn_Jet.Execute sqlDoc
        
        '결과 표시
        Set itemH = lvwComplete.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
        If itemH Is Nothing Then
            Set itemH = lvwComplete.ListItems.Add()
            With itemH
                .Key = COL_KEY & Result.Rst_Sid '아이템 키에 검체번호
                .Text = Result.Rst_Sid          '아이템 에 검체번호
                .Tag = Result.Rst_Type          '테그에 결과 타입
                .SmallIcon = "ITM" '"ITM"
            End With
        End If
        '결과값 등록
        itemH.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
        itemH.ForeColor = vbBlack
        If Result.Rst_Eid = "1" Then itemH.ForeColor = vbRed
        
        Set itemS = itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
        
        itemS.Tag = Result.Rst_Error '서브아이템 테그에 에러 메시지
        Set itemS = Nothing
        Set itemX = Nothing
        Set itemX = Nothing
    End If
    '검사코드가 없는것은 등록 하지 않음
    Exit Sub
ErrorRoutine:

    Set itemS = Nothing
    Set itemX = Nothing
    Set itemX = Nothing
    
    Call ErrMsgProc(CallForm)

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


