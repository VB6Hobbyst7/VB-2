VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmStatistics 
   Caption         =   "��� �Է�"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15285
   WindowState     =   2  '�ִ�ȭ
   Begin TabDlg.SSTab SSTab1 
      Height          =   8445
      Left            =   60
      TabIndex        =   6
      Top             =   450
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14896
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "���"
      TabPicture(0)   =   "frmStatistics.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tblexcel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CommonDialog1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "sspTest"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "sspDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSerch"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optCondition(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optCondition(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lvwCuData"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkQC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "spdSumS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "spdStaTotal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "spdStatistics"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "���� ����"
      TabPicture(1)   =   "frmStatistics.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "spdKitCode"
      Tab(1).Control(2)=   "spdSugaSet"
      Tab(1).ControlCount=   3
      Begin FPSpread.vaSpread spdStatistics 
         Height          =   7485
         Left            =   150
         TabIndex        =   28
         Top             =   870
         Width           =   10035
         _Version        =   196608
         _ExtentX        =   17701
         _ExtentY        =   13203
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmStatistics.frx":0038
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdStaTotal 
         Height          =   7485
         Left            =   10230
         TabIndex        =   30
         Top             =   870
         Width           =   4815
         _Version        =   196608
         _ExtentX        =   8493
         _ExtentY        =   13203
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   3
         MaxRows         =   5
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmStatistics.frx":05EA
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdSumS 
         Height          =   1215
         Left            =   600
         TabIndex        =   35
         Top             =   7110
         Visible         =   0   'False
         Width           =   14445
         _Version        =   196608
         _ExtentX        =   25479
         _ExtentY        =   2143
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   27
         MaxRows         =   2
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmStatistics.frx":0A23
         UserResize      =   0
      End
      Begin VB.CheckBox chkQC 
         Caption         =   "QC����"
         Height          =   315
         Left            =   6330
         TabIndex        =   33
         Top             =   480
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4860
         Left            =   5190
         TabIndex        =   29
         Top             =   2700
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   8573
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
      Begin VB.OptionButton optCondition 
         Caption         =   "�˻��׸�"
         Height          =   285
         Index           =   1
         Left            =   10290
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optCondition 
         Caption         =   "�Ⱓ��"
         Height          =   285
         Index           =   0
         Left            =   9330
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Caption         =   " ���� ���� "
         Height          =   2565
         Left            =   -65340
         TabIndex        =   7
         Top             =   690
         Width           =   3225
         Begin VB.CommandButton cmdSDelete 
            Caption         =   "����"
            Height          =   375
            Left            =   2130
            TabIndex        =   11
            Top             =   1950
            Width           =   855
         End
         Begin VB.CommandButton cmdSSave 
            Caption         =   "����"
            Height          =   375
            Left            =   1170
            TabIndex        =   10
            Top             =   1950
            Width           =   855
         End
         Begin VB.TextBox txtSuga 
            Height          =   345
            Left            =   1200
            TabIndex        =   9
            Top             =   1380
            Width           =   1785
         End
         Begin VB.TextBox txtSugaCnt 
            Height          =   345
            Left            =   1200
            TabIndex        =   8
            Top             =   900
            Width           =   1785
         End
         Begin VB.Label Label12 
            Caption         =   "�˻��"
            Height          =   255
            Left            =   330
            TabIndex        =   15
            Top             =   975
            Width           =   555
         End
         Begin VB.Label Label11 
            Caption         =   "���� "
            Height          =   255
            Left            =   330
            TabIndex        =   14
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblKitCode 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BorderStyle     =   1  '���� ����
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Top             =   450
            Width           =   1785
         End
         Begin VB.Label Label5 
            Caption         =   "KIT �ڵ�"
            Height          =   255
            Left            =   330
            TabIndex        =   12
            Top             =   510
            Width           =   765
         End
      End
      Begin BHButton.BHImageButton cmdSerch 
         Height          =   360
         Left            =   4980
         TabIndex        =   16
         Top             =   450
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         Caption         =   "��ȸ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread spdKitCode 
         Height          =   6465
         Left            =   -74520
         TabIndex        =   19
         Top             =   780
         Width           =   2925
         _Version        =   196608
         _ExtentX        =   5159
         _ExtentY        =   11404
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   1
         MaxRows         =   1
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735309
         SpreadDesigner  =   "frmStatistics.frx":0FBB
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdSugaSet 
         Height          =   6465
         Left            =   -71250
         TabIndex        =   20
         Top             =   780
         Width           =   5565
         _Version        =   196608
         _ExtentX        =   9816
         _ExtentY        =   11404
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   3
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmStatistics.frx":1269
         UserResize      =   0
      End
      Begin Threed.SSPanel sspDate 
         Height          =   435
         Left            =   150
         TabIndex        =   21
         Top             =   390
         Width           =   4755
         _Version        =   65536
         _ExtentX        =   8387
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   300
            Left            =   1350
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   60
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   94961665
            CurrentDate     =   37112
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   300
            Left            =   3060
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   94961665
            CurrentDate     =   37112
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '����
            Caption         =   "�۾����� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel sspTest 
         Height          =   435
         Left            =   150
         TabIndex        =   25
         Top             =   390
         Width           =   4755
         _Version        =   65536
         _ExtentX        =   8387
         _ExtentY        =   767
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.ComboBox cboTest 
            Height          =   300
            ItemData        =   "frmStatistics.frx":15CF
            Left            =   1350
            List            =   "frmStatistics.frx":15D1
            TabIndex        =   27
            Top             =   60
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '����
            Caption         =   "�˻�� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   240
            TabIndex        =   26
            Top             =   120
            Width           =   1095
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   360
         Left            =   7680
         TabIndex        =   31
         Top             =   420
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         Caption         =   "Excel ���� ����"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread tblexcel 
         Height          =   675
         Left            =   3030
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   675
         _Version        =   196608
         _ExtentX        =   1191
         _ExtentY        =   1191
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
         SpreadDesigner  =   "frmStatistics.frx":15D3
      End
      Begin VB.Label Label8 
         Caption         =   "�޾缺��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   34
         Top             =   7170
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   10770
      Top             =   30
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
            Picture         =   "frmStatistics.frx":177E
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":1D18
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":22B2
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":284C
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":2DE6
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":3380
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   0
      Top             =   8790
      Width           =   11940
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1410
         TabIndex        =   3
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Delete"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   2730
         TabIndex        =   4
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4050
         TabIndex        =   5
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '�� ����
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   714
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmStatistics.frx":391A
      Caption         =   " �˻� ���"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mAdoRs As ADODB.Recordset

Private Const COL_WIDTH As Long = "900"

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "����"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "��Ϲ�ȣ"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "��  ��"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "��ü��ȣ"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "��ü��ȣ"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "�� ��"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "�˻��׸�"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '��� �ڵ�
Private Const TEST_CD_LIS   As String = "LIS_CD"    '�˻�� �ڵ�
Private Const TEST_NM_LIS   As String = "LIS_NM"    '�˻�� �̸�
Private Const TEST_VALUES   As String = "VALUES"    '���


Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdSave_Click
        Case 1
            Call cmdPrint2_Click
        Case 2
            Call cmdClear_Click
        Case 3 'cmd close
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub cmdPrint2_Click()

End Sub

Private Sub cmdClear_Click()
    'ClearSpread spdKitCode
    ClearSpread spdSugaSet
    ClearSpread spdSumS
    
    lblKitCode.Caption = ""
    txtSugaCnt.text = ""
    txtSuga.text = ""
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Progress_View()
Dim i As Long

    Call SetProgress(1005, Custom, "Loding", True)
    
    For i = 1 To 1001
        Call ShowProgress(i, "TEST " & i, True)
    Next
    Call SetProgress(100, Custom, "End", False)
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim strTmp1 As String
    Dim lngRows As Long
    
    If spdStatistics.DataRowCnt = 0 And spdStatistics.DataRowCnt = 0 Then Exit Sub
    
    With spdStatistics
        .Row = 0: .Row2 = .maxrows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .maxrows
    End With
 
    With spdStaTotal
        .Row = 0: .Row2 = .maxrows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp1 = .Clip
        .BlockMode = False
        lngRows = .maxrows
    End With
 
'    With tblexcel
'        .maxrows = spdStatistics.maxrows + 1
'        .MaxCols = spdStatistics.MaxCols
'        .Row = 1: .Row2 = .maxrows
'        .Col = 1: .Col2 = spdStatistics.MaxCols
'        .BlockMode = True
'        .Clip = strTmp
'        .BlockMode = False
'    End With
    
    With tblexcel
        .maxrows = spdStatistics.maxrows + 10
        .MaxCols = spdStatistics.MaxCols
        .Row = 1: .Row2 = .maxrows
        .Col = 1: .Col2 = spdStatistics.MaxCols
        .BlockMode = True
        .Clip = strTmp & vbNewLine & strTmp1
        .BlockMode = False
    End With
'
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.filter = "ExCelFile(*.XLS)|*.XLS"
    CommonDialog1.FileName = REG_INSNAME & "  " & Format(dtpToDate, "yyyy-mm-dd") & " �˻���Ȳ����"
    CommonDialog1.ShowSave

    tblexcel.SaveTabFile (CommonDialog1.FileName)

End Sub

Private Sub cmdSDelete_Click()
    Dim strSugaCnt      As String
    Dim strKitCode      As String
    Dim i As Integer
    Dim strSql          As String
    
    strKitCode = Trim(lblKitCode.Caption)
    strSugaCnt = Trim(txtSugaCnt)
    
    If strKitCode = "" Then
        Exit Sub
    End If
    
    If strSugaCnt = "" Then
        txtSugaCnt.SetFocus
        Exit Sub
    End If
    
    
             strSql = "DELETE FROM EQUIPSUGA "
    strSql = strSql & " WHERE EQP_CD = '" & INS_CODE & "' AND KITCODE = '" & strKitCode & "' AND EXAMCNT = '" & strSugaCnt & "'"
    AdoCn_Jet.Execute strSql
        
    For i = 1 To spdKitCode.DataRowCnt
        If Trim(GetText(spdKitCode, i, 1)) = strKitCode Then
            spdKitCode_Click 1, i
            Exit For
        End If
    Next

End Sub

Private Sub cmdSerch_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim adoRS1  As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim IntRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    Dim strRackNo, strPos As String
    
    Dim itemX       As ListItem
    Dim intCnt      As Long
    
    Dim strKitCode As String
    Dim pGrid_Point As Integer
    Dim strVal0, strVal1, strVal2, strVal3

    Dim i As Integer
    Dim varTmp As Variant
    
    IntRow = 0
    intCnt = 0
    
    With spdStaTotal
        .maxrows = 1
        .MaxCols = 3
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdStatistics
        .maxrows = 1
'        .MaxCols = 12
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdSumS
        .maxrows = 2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    strVal0 = 0
    
    sqlDoc = "Select SPCNO, PATID, S_NO1, TESTCD, EQUIPCD, EQPNUM, TRANSDT, RSTVAL, REFVAL, TRANSDT,TRANSTM, EQPNUM, PATID, PNM, SEX " & _
             "  From INTERFACE003" & _
             " Where TRANSDT BetWeen '" & Format(dtpFromDate.Value, "yyyymmdd") & "' And '" & Format(dtpToDate.Value, "yyyymmdd") & "'" & _
             "   And EQUIPCD = '" & INS_CODE & "'" & _
             "   And (S_NO1 <> '' or S_NO1 is not null ) "
    If chkQC.Value <> "1" Then
        sqlDoc = sqlDoc & " And PNM <> 'QC' "
    End If
    sqlDoc = sqlDoc & " Order By TRANSDT, SPCNO, S_NO1,TESTCD "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdStatistics
            intCnt = intCnt + 1
            If strSpcno <> Trim$(adoRS("TRANSDT") & "") + Trim$(adoRS("SPCNO") & "") + Trim$(adoRS("S_NO1") & "") Then
                IntRow = IntRow + 1
                intCnt = 1
                If IntRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'                .SetText 1, IntRow, Format(Trim$(adoRS("TRANSDT")), "##-##-##") & " " & Mid(Trim$(adoRS("TRANSTM")), 1, 2) & ":" & Mid(Trim$(adoRS("TRANSTM")), 3, 2)
                .SetText 1, IntRow, Format(Trim$(adoRS("TRANSDT")), "##-##-##") ' & " " & Format(Trim$(adoRS("TRANSTM")), "##:##:##")
                .SetText 2, IntRow, IIf(Trim$(adoRS("PATID") & "") = "", "QC", Trim$(adoRS("PATID") & ""))
                .SetText 3, IntRow, Trim$(adoRS("EQPNUM") & "")
                .SetText 4, IntRow, Trim$(adoRS("PATID") & "")
                .SetText 5, IntRow, Trim$(adoRS("PNM") & "")
                .SetText 6, IntRow, Trim$(adoRS("S_NO1") & "")
                .SetText 7, IntRow, intCnt
                         
                         sqlDoc = "SELECT SUGA FROM EQUIPSUGA "
                sqlDoc = sqlDoc & " WHERE EQP_CD = '" & INS_CODE & "'"
                sqlDoc = sqlDoc & "   AND KITCODE = '" & Trim$(adoRS("S_NO1") & "") & "'"
                sqlDoc = sqlDoc & "   AND EXAMCNT = '" & intCnt & "'"
                
                Set adoRS1 = New ADODB.Recordset
                adoRS1.CursorLocation = adUseClient
                adoRS1.Open sqlDoc, AdoCn_Jet
                If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
                Do While Not adoRS1.EOF
                    strVal0 = Trim$(adoRS1("SUGA") & "")
                    adoRS1.MoveNext
                Loop
                adoRS1.Close:    Set adoRS1 = Nothing
                
                .SetText 8, IntRow, strVal0
                
                'If IntRow > 1 Then intCnt = 0
                
                For i = 9 To .MaxCols
                    .GetText i, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        If Trim$(adoRS("TESTCD") & "") <> "" And Trim$(adoRS("TESTCD") & "") = itemX.tag Then
                            'blnSameCode = False
                            .SetText i, IntRow, Trim$(adoRS("TESTCD") & "")
                            Exit For
                        End If
                    End If
                    Set itemX = Nothing
                Next
                spdStaTotal.Row = 1
                
            Else
                         sqlDoc = "SELECT SUGA FROM EQUIPSUGA "
                sqlDoc = sqlDoc & " WHERE EQP_CD = '" & INS_CODE & "'"
                sqlDoc = sqlDoc & "   AND KITCODE = '" & Trim$(adoRS("S_NO1") & "") & "'"
                sqlDoc = sqlDoc & "   AND EXAMCNT = '" & intCnt & "'"

                Set adoRS1 = New ADODB.Recordset
                adoRS1.CursorLocation = adUseClient
                adoRS1.Open sqlDoc, AdoCn_Jet
                If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
                Do While Not adoRS1.EOF
                    strVal0 = Trim$(adoRS1("SUGA") & "")
                    adoRS1.MoveNext
                Loop
                adoRS1.Close:    Set adoRS1 = Nothing
                        
                .SetText 7, IntRow, intCnt
                .SetText 8, IntRow, strVal0
                strVal0 = 0
                
                For i = 9 To .MaxCols
                    .GetText i, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        If Trim$(adoRS("TESTCD") & "") <> "" And Trim$(adoRS("TESTCD") & "") = itemX.tag Then
                            'blnSameCode = False
                            .SetText i, IntRow, Trim$(adoRS("TESTCD") & "")
                            Exit For
                        End If
                    End If
                    Set itemX = Nothing
                Next
                
                'intCnt = 0
            End If
            strSpcno = Trim$(adoRS("TRANSDT") & "") + Trim$(adoRS("SPCNO") & "") + Trim$(adoRS("S_NO1") & "")
        End With
        adoRS.MoveNext
    Loop
    
    adoRS.Close:    Set adoRS = Nothing


On Error Resume Next

    strVal2 = 0
    strVal3 = 0
    
    With spdStaTotal
        For intCnt = 1 To spdStatistics.maxrows
            spdStatistics.Row = intCnt
            spdStatistics.Col = 6
            strVal1 = spdStatistics.text
            
            pGrid_Point = SeqSearch(spdStaTotal, spdStatistics.text, 1)

            If pGrid_Point = 0 Then
                pGrid_Point = SeqNullSearch(spdStaTotal, spdStaTotal.text, 1)
                If pGrid_Point = 0 Then
                    spdStaTotal.maxrows = spdStaTotal.maxrows + 1: pGrid_Point = spdStaTotal.maxrows
                    spdStaTotal.RowHeight(-1) = 12
                End If
            End If
            
            spdStatistics.Row = intCnt
            spdStatistics.Col = 6: strVal1 = spdStatistics.text
            spdStatistics.Col = 7: strVal2 = spdStatistics.text
            
                     sqlDoc = "SELECT SUGA FROM EQUIPSUGA "
            sqlDoc = sqlDoc & " WHERE EQP_CD = '" & INS_CODE & "'"
            sqlDoc = sqlDoc & "   AND KITCODE = '" & strVal1 & "'"
            sqlDoc = sqlDoc & "   AND EXAMCNT = '" & strVal2 & "'"
            
            Set adoRS = New ADODB.Recordset
            adoRS.CursorLocation = adUseClient
            adoRS.Open sqlDoc, AdoCn_Jet
            If adoRS.RecordCount > 0 Then adoRS.MoveFirst
            Do While Not adoRS.EOF
                strVal3 = Trim$(adoRS("SUGA") & "")
                adoRS.MoveNext
            Loop
            adoRS.Close:    Set adoRS = Nothing
            
 '           spdStatistics.Col = 11: strVal2 = CLng(spdStatistics.text)
 '           spdStatistics.Col = 12: strVal3 = CLng(strVal3) + CLng(spdStatistics.text)
            
            spdStaTotal.Row = pGrid_Point
            spdStaTotal.Col = 2: strVal2 = CLng(strVal2) + CLng(spdStaTotal.text)
            spdStaTotal.Col = 3: strVal3 = CLng(strVal3) + CLng(spdStaTotal.text)
            
            
            spdStaTotal.SetText 1, pGrid_Point, strVal1
            spdStaTotal.SetText 2, pGrid_Point, strVal2
            spdStaTotal.SetText 3, pGrid_Point, strVal3
            
            strVal2 = 0
            strVal3 = 0
        Next
    End With
    
    '-- �缺��
    Dim varSum(40) As Long
    Dim varPSum(40) As Long
    
    sqlDoc = "Select TESTCD, RSTVAL " & _
             "  From INTERFACE003" & _
             " Where TRANSDT BetWeen '" & Format(dtpFromDate.Value, "yyyymmdd") & "' And '" & Format(dtpToDate.Value, "yyyymmdd") & "'" & _
             "   And EQUIPCD = '" & INS_CODE & "'" & _
             "   And (S_NO1 <> '' or S_NO1 is not null ) "
    If chkQC.Value <> "1" Then
        sqlDoc = sqlDoc & " And PATID <> '' "
    End If
    sqlDoc = sqlDoc & " Order By TESTCD "
    
    
'        Set adoRS2 = New ADODB.Recordset
'
'                 sqlDoc = "SELECT TMP1 " & vbCrLf
'        sqlDoc = sqlDoc & "  From INTERFACE003 " & vbCrLf
'        sqlDoc = sqlDoc & " Where SPCNO = '" & strBarno & "'"
'        sqlDoc = sqlDoc & "   And TESTCD = '" & varTmp & "'"
'        sqlDoc = sqlDoc & "   And TRANSDT = '" & Format(Now, "yyyymmdd") & "'"
'        sqlDoc = sqlDoc & "   And TMP1 is not null"
'
'        adoRS2.CursorLocation = adUseClient
'        adoRS2.Open sqlDoc, AdoCn_Jet
'        If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
'        Do While Not adoRS2.EOF
'            strTestCd = adoRS2.Fields("TMP1").Value & ""
'            adoRS2.MoveNext
'        Loop
'        adoRS2.Close:    Set adoRS2 = Nothing
    
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    With spdSumS
        Do Until adoRS.EOF
            For i = 1 To .MaxCols
                .GetText i, 0, varTmp
                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                If Not itemX Is Nothing Then
                    If Trim$(adoRS("TESTCD") & "") <> "" And Trim$(adoRS("TESTCD") & "") = itemX.tag Then
                        .GetText i, IntRow, varSum(i)
                        varSum(i) = varSum(i) + 1
                        
                        
                        If Trim$(adoRS("RSTVAL") & "") = "Positive" Then
                            varPSum(i) = varPSum(i) + 1
                        End If      'Format((cntPos / cntSum) * 100, "#0.00")
                        .SetText i, 1, Format((varPSum(i) / varSum(i)) * 100, "#0.00")
                        .SetText i, 2, varPSum(i) & "/" & varSum(i)
                    Else
                        .GetText i, IntRow, varSum(i)
                        varSum(i) = varSum(i) + 0
                        If varPSum(i) = 0 Then
                            .SetText i, 1, "0.00"
                        Else
                            .SetText i, 1, Format((varPSum(i) / varSum(i)) * 100, "#0.00")
                        End If
                        .SetText i, 2, varPSum(i) & "/" & varSum(i)
                        
                    End If
                End If
                Set itemX = Nothing
            Next
            adoRS.MoveNext
        Loop
    End With
End Sub

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If Trim(.text) = "" Then
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
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If .text = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

'
'Private Sub setStaTotal()
'    Dim intCnt As Integer
'    Dim i As Integer
'    Dim blnSame As Boolean
'
'    blnSame = False
'
'    With spdStaTotal
'        For i = 1 To .maxrows
'            If .text = "" Then
'                .Row = i
'                .SetText 1, .maxrows, strKitCode
'                .SetText 2, .maxrows, strCnt
'                .SetText 3, .maxrows, strValue
'                blnSame = True
'                Exit For
'
'            Else
'                .Col = 1
'                If .text = strKitCode Then
'                    .Col = 2:   .text = .text + strCnt
'                    .Col = 3:   .text = .text + strValue
'                    blnSame = True
'                    Exit For
'                End If
'            End If
'        Next
'
'        If blnSame = False Then
'            .maxrows = .maxrows + 1
'            .SetText 1, .maxrows, strKitCode
'            .SetText 2, .maxrows, strCnt
'            .SetText 3, .maxrows, strValue
'        End If
'
'    End With
'
'
'End Sub

Private Sub cmdSSave_Click()
    Dim strSugaCnt      As String
    Dim strSuga         As String
    Dim strKitCode      As String
    Dim i As Integer
    Dim strSql          As String
    
    strKitCode = Trim(lblKitCode.Caption)
    strSugaCnt = Trim(txtSugaCnt)
    strSuga = Trim(txtSuga)
    
    If strKitCode = "" Then
        Exit Sub
    End If
    
    If strSugaCnt = "" Then
        txtSugaCnt.SetFocus
        Exit Sub
    End If
    
    If strSuga = "" Then
        txtSuga.SetFocus
        Exit Sub
    End If
    
             strSql = "DELETE FROM EQUIPSUGA "
    strSql = strSql & " WHERE EQP_CD = '" & INS_CODE & "' AND KITCODE = '" & strKitCode & "' AND EXAMCNT = '" & strSugaCnt & "'"
    AdoCn_Jet.Execute strSql
    
             strSql = "INSERT INTO EQUIPSUGA "
    strSql = strSql & " (EQP_CD,KITCODE,EXAMCNT,SUGA) "
    strSql = strSql & " VALUES "
    strSql = strSql & " ('" & INS_CODE & "', '" & strKitCode & "','" & strSugaCnt & "','" & strSuga & "')"
    AdoCn_Jet.Execute strSql
    
    For i = 1 To spdKitCode.DataRowCnt
        If Trim(GetText(spdKitCode, i, 1)) = strKitCode Then
            spdKitCode_Click 1, i
            Exit For
        End If
    Next

End Sub

'Private Function Drow_Header() As Boolean
'    Dim objStatics      As clsStatistics
'    Dim itemKey         As String
'    Dim itemText        As String
'    Dim AdoRs_TstNm     As ADODB.Recordset
'
'    Set objStatics = New clsStatistics
'    Drow_Header = True
'    With objStatics
'        .SetAdoCn AdoCn_Jet
'
'        Set AdoRs_TstNm = .Get_TestName(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
'        If Not AdoRs_TstNm Is Nothing Then
'            If AdoRs_TstNm.EOF Then
'                Drow_Header = False
'            Else
'                lvwStatics(0).ColumnHeaders.Clear
'                lvwStatics(0).ColumnHeaders.Add , "DATE", "DATE"
'                AdoRs_TstNm.MoveFirst
'                Do Until AdoRs_TstNm.EOF
'                    itemKey = Trim(AdoRs_TstNm.Fields("TESTCD") & "")
'                    itemText = Trim(AdoRs_TstNm.Fields("TESTNM") & "")
'                    Call lvwStatics(0).ColumnHeaders.Add(, itemKey, itemText, COL_WIDTH, lvwColumnRight)
'                    AdoRs_TstNm.MoveNext
'                Loop
'                lvwStatics(0).HideColumnHeaders = False
'            End If
'        Else
'            Drow_Header = False
'        End If
'    End With
'
'    Set AdoRs_TstNm = Nothing
'    Set objStatics = Nothing
'End Function
'
'Private Function Drow_Date(ByVal Index As Integer) As Boolean
'    Dim objStatics      As clsStatistics
'    Dim itemKey         As String
'    Dim itemText        As String
'    Dim AdoRs_TstDt     As ADODB.Recordset
'
'    Set objStatics = New clsStatistics
'    Drow_Date = True
'    With objStatics
'        .SetAdoCn AdoCn_Jet
'        Set AdoRs_TstDt = .Get_TestDate(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
'        If Not AdoRs_TstDt Is Nothing Then
'            If AdoRs_TstDt.EOF Then
'                Drow_Date = False
'            Else
'                AdoRs_TstDt.MoveFirst
'                Do Until AdoRs_TstDt.EOF
'                    itemKey = Trim(AdoRs_TstDt.Fields("ACCDT") & "")
'                    itemText = Trim(AdoRs_TstDt.Fields("ACCDT") & "")
'                    lvwStatics(Index).ListItems.Add , itemKey, itemText, , "LST"
'                    AdoRs_TstDt.MoveNext
'                Loop
'            End If
'        Else
'            Drow_Date = False
'        End If
'    End With
'
'    Set AdoRs_TstDt = Nothing
'    Set objStatics = Nothing
'End Function
'
'Private Function Drow_Item() As Boolean
'    Dim objStatics      As clsStatistics
'    Dim itemKey         As String
'    Dim itemHeadKey     As String
'    Dim itemText        As String
'    Dim AdoRs_TstCn     As ADODB.Recordset
'
'    Set objStatics = New clsStatistics
'    Drow_Item = True
'    With objStatics
'        .SetAdoCn AdoCn_Jet
'        Set AdoRs_TstCn = .Get_TestCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
'        If Not AdoRs_TstCn Is Nothing Then
'            If AdoRs_TstCn.EOF Then
'                Drow_Item = False
'            Else
'                Do Until AdoRs_TstCn.EOF
'                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
'                    itemHeadKey = Trim(AdoRs_TstCn.Fields("TESTCD") & "")
'                    itemText = Trim(AdoRs_TstCn.Fields("CNT") & "")
'                    lvwStatics(0).ListItems(itemKey).SubItems(lvwStatics(0).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
'                    AdoRs_TstCn.MoveNext
'                Loop
'            End If
'        Else
'            Drow_Item = False
'        End If
'
'    End With
'    Set AdoRs_TstCn = Nothing
'    Set objStatics = Nothing
'End Function
'
'Private Sub Total_Calculation(ByVal Index As Integer)
'    Dim itemX           As ListItem
'    Dim itemS           As ListSubItem
'    Dim lngTotal()      As Long
'    Dim i As Long
'
'    ReDim lngTotal(lvwStatics(Index).ColumnHeaders.count - 1)
'    For Each itemX In lvwStatics(Index).ListItems
'        For i = 1 To lvwStatics(Index).ColumnHeaders.count - 1
'            lngTotal(i) = lngTotal(i) + Val(itemX.SubItems(i))
'        Next
'    Next
'    Set itemX = lvwStatics(Index).ListItems.Add
'    With itemX
'        .text = "TOTAL"
'        .Bold = True
'    End With
'
'    For i = 1 To lvwStatics(Index).ColumnHeaders.count - 1
'        Set itemS = itemX.ListSubItems.Add(i)
'        With itemS
'            .Bold = True
'            .ForeColor = vbBlue
'            .text = lngTotal(i)
'        End With
'    Next
'
'    Set itemX = Nothing
'
'End Sub
'
'Private Function Drow_SlipCount() As Boolean
'    Dim objStatics      As clsStatistics
'    Dim itemKey         As String
'    Dim itemHeadKey     As String
'    Dim itemText        As String
'    Dim AdoRs_TstCn     As ADODB.Recordset
'
'    Set objStatics = New clsStatistics
'    Drow_SlipCount = True
'    With objStatics
'        .SetAdoCn AdoCn_Jet
'        Set AdoRs_TstCn = .Get_SlipCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
'        If Not AdoRs_TstCn Is Nothing Then
'            If AdoRs_TstCn.EOF Then
'                Drow_SlipCount = False
'            Else
'                Do Until AdoRs_TstCn.EOF
'                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
'                    itemHeadKey = "KEY_SLIP"
'                    itemText = Trim(AdoRs_TstCn.Fields("SLIP_CNT") & "")
'                    lvwStatics(1).ListItems(itemKey).SubItems(lvwStatics(1).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
'                    AdoRs_TstCn.MoveNext
'                Loop
'            End If
'        Else
'            Drow_SlipCount = False
'        End If
'
'    End With
'    Set AdoRs_TstCn = Nothing
'    Set objStatics = Nothing
'End Function
'
'Private Function Drow_TestCount() As Boolean
'    Dim objStatics      As clsStatistics
'    Dim itemKey         As String
'    Dim itemHeadKey     As String
'    Dim itemText        As String
'    Dim AdoRs_TstCn     As ADODB.Recordset
'
'    Set objStatics = New clsStatistics
'    Drow_TestCount = True
'    With objStatics
'        .SetAdoCn AdoCn_Jet
'        Set AdoRs_TstCn = .Get_TotalTestCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
'        If Not AdoRs_TstCn Is Nothing Then
'            If AdoRs_TstCn.EOF Then
'                Drow_TestCount = False
'            Else
'                Do Until AdoRs_TstCn.EOF
'                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
'                    itemHeadKey = "KEY_TEST"
'                    itemText = Trim(AdoRs_TstCn.Fields("TEST_CNT") & "")
'                    lvwStatics(1).ListItems(itemKey).SubItems(lvwStatics(1).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
'                    AdoRs_TstCn.MoveNext
'                Loop
'            End If
'        Else
'            Drow_TestCount = False
'        End If
'
'    End With
'    Set AdoRs_TstCn = Nothing
'    Set objStatics = Nothing
'End Function

Private Sub Form_Load()
    
    dtpFromDate.Value = Now - 7
    dtpToDate.Value = Now
    optCondition(0).Value = True
    
    spdKitCode.maxrows = 1
    spdSugaSet.maxrows = 1
    
    Call DisplayKITList
    
    Call DisplayTestList
    
    Call f_subSet_StatHeader
    Call f_subSet_StatList

    SSTab1.Tab = 0
    
End Sub



Private Sub f_subSet_StatHeader()
    
    '�˻��ڵ� ���̺�
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
            Call .Add(, TEST_CD_LIS, "�˻��ڵ�", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "�� �� ��", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "�˻���", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFL", "����ġ(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "����ġ(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "���", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "��ü�ڵ�", (lvwCuData.Width - 310) * 0.5)
            Call .Add(, "TESTNO", "KIT�ڵ�", (lvwCuData.Width - 310) * 0.5)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub


Private Sub f_subSet_StatList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, IntRow  As Integer
    Dim intsCol As Integer
    Dim intPos1 As Integer
    
    Dim mIntNms As clsIntTest


'On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Private Sub f_subSet_StatList()"
    
    intsCol = 1
    intCol = 10
    intCol2 = 1
    IntRow = 1
    
    lvwCuData.ListItems.Clear
    
    With spdStatistics
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH,TESTNO" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
    End If
    
    Do While Not adoRS.EOF
        
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM_EQP") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.SubItems(12) = Trim(adoRS.Fields("TESTNO") & "")
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
'            itemX.text = Trim(adoRS.Fields("TESTNO") & "")
        Set itemX = Nothing
        
        With spdStatistics
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
            End If
            .ColWidth(intCol) = 4.5
            .SetText intCol, 0, Trim$(adoRS("TESTNM_EQP") & "")
        End With
        
        With spdSumS
            If intsCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
            End If
            .ColWidth(intsCol) = 4.5
            .SetText intsCol, 0, Trim$(adoRS("TESTNM_EQP") & "")
        End With
        
'        fChannel(intCol - 8) = adoRS.Fields("TEST_EQP")
        
        intsCol = intsCol + 1
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    
    Set adoRS = Nothing
    

Exit Sub

ErrRoutine:
    Set adoRS = Nothing
'    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub DisplayKITList()
    
    Dim objStatics      As clsStatistics
    Dim AdoRs_TstCn     As ADODB.Recordset
    Dim IntRow          As Integer
    
    Call ClearSpread(spdKitCode)
    Call ClearSpread(spdStaTotal)
    IntRow = 0
    
    Set objStatics = New clsStatistics
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_KitCode
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
            
            Else
                spdKitCode.maxrows = AdoRs_TstCn.RecordCount
                Do Until AdoRs_TstCn.EOF
                    IntRow = IntRow + 1
                    Call spdKitCode.SetText(1, IntRow, Trim(AdoRs_TstCn.Fields("TESTNO").Value & ""))
                    AdoRs_TstCn.MoveNext
                Loop
                spdKitCode.RowHeight(-1) = 12
            End If
        End If
        
    End With
    
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
    
End Sub

Private Sub DisplayTestList()
    
    Dim objStatics      As clsStatistics
    Dim AdoRs_TstCn     As ADODB.Recordset
'    Dim IntRow          As Integer
    Dim strK1           As String
    Dim strK2           As String
    
    Call ClearSpread(spdStatistics)
'    IntRow = 0
    
    Set objStatics = New clsStatistics
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_TestCode
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
            
            Else
                'spdKitCode.maxrows = AdoRs_TstCn.RecordCount
                cboTest.AddItem "== ��ü =="
                Do Until AdoRs_TstCn.EOF
                    cboTest.AddItem AdoRs_TstCn.Fields("TESTNM").Value   'TESTNO, TESTCD_EQP, TESTNM
                    AdoRs_TstCn.MoveNext
                Loop
            End If
        End If
        
    End With
    
    cboTest.ListIndex = 0
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
    
End Sub


Private Sub DisplaySUGAList(ByVal varKITCode As Variant)
    
    Dim objStatics      As clsStatistics
    Dim AdoRs_TstCn     As ADODB.Recordset
    Dim IntRow          As Integer
    
    Call ClearSpread(spdSugaSet)
    IntRow = 0
    
    Set objStatics = New clsStatistics
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_SugaCode(varKITCode)
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
            
            Else
                spdSugaSet.maxrows = AdoRs_TstCn.RecordCount
                Do Until AdoRs_TstCn.EOF
                    IntRow = IntRow + 1
                    Call spdSugaSet.SetText(1, IntRow, Trim(AdoRs_TstCn.Fields("KITCODE").Value & ""))
                    Call spdSugaSet.SetText(2, IntRow, Trim(AdoRs_TstCn.Fields("EXAMCNT").Value & ""))
                    Call spdSugaSet.SetText(3, IntRow, Trim(AdoRs_TstCn.Fields("SUGA").Value & ""))
                    AdoRs_TstCn.MoveNext
                Loop
                spdSugaSet.RowHeight(-1) = 12
            End If
        End If
        
    End With
    
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
    
End Sub

Private Sub Form_Resize()
    
    
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub



Private Sub optCondition_Click(Index As Integer)

    If Index = 0 Then
        sspDate.ZOrder 0
    Else
        sspTest.ZOrder 0
    End If
End Sub

Private Sub spdKitCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim varKITCode As Variant
    
    Call cmdClear_Click
    
    spdKitCode.GetText 1, Row, varKITCode

    DisplaySUGAList varKITCode
    
    
    
    lblKitCode.Caption = varKITCode
    
End Sub

Private Sub spdSugaSet_Click(ByVal Col As Long, ByVal Row As Long)
    
    With spdSugaSet
        .Row = Row
        .Col = 1: lblKitCode.Caption = Trim(.text)
        .Col = 2: txtSugaCnt.text = Trim(.text)
        .Col = 3: txtSuga.text = Trim(.text)
    End With

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Dim sNo As String

    If PreviousTab = 0 Then
        sNo = InputBox("��й�ȣ�� �Է��ϼ��� !")
        If sNo <> Format(Now, "yyyymmdd") Then
          SSTab1.Tab = 0
        End If
    End If
    
End Sub

