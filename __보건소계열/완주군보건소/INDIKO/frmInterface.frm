VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BackColor       =   &H00BF8B59&
   Caption         =   "SANSOFT"
   ClientHeight    =   10290
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   22260
   BeginProperty Font 
      Name            =   "����"
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
   ScaleHeight     =   10290
   ScaleWidth      =   22260
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox picComm 
      Align           =   2  '�Ʒ� ����
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   22200
      TabIndex        =   7
      Top             =   9030
      Visible         =   0   'False
      Width           =   22260
      Begin VB.CommandButton cmdRcvView 
         Caption         =   "�α׺���"
         Height          =   435
         Left            =   21480
         TabIndex        =   23
         Top             =   90
         Width           =   945
      End
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "�����"
         Height          =   525
         Left            =   12930
         TabIndex        =   17
         Top             =   60
         Width           =   885
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   16
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   15
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   14
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   13
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   12
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         BackColor       =   &H00C0FFFF&
         Height          =   555
         Left            =   13950
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   60
         Width           =   3045
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "������"
         Height          =   525
         Left            =   17010
         TabIndex        =   10
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "�ޱ�"
         Height          =   525
         Left            =   11940
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraHidden 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hidden"
      Height          =   3075
      Left            =   7590
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   5175
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
         Height          =   705
         Left            =   60
         TabIndex        =   64
         Top             =   2100
         Width           =   3945
         Begin VB.TextBox txtFrNo 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1770
            TabIndex        =   66
            Text            =   "0000"
            Top             =   180
            Width           =   765
         End
         Begin VB.TextBox txtToNo 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2820
            TabIndex        =   65
            Text            =   "0999"
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label1 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "~"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   2610
            TabIndex        =   70
            Top             =   270
            Width           =   150
         End
         Begin VB.Label lblSlipCd 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            BorderStyle     =   1  '���� ����
            Caption         =   "L20"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   69
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label3 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            BorderStyle     =   1  '���� ����
            Caption         =   "00I"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   68
            Top             =   210
            Width           =   465
         End
         Begin VB.Label Label1 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "SLIP"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   67
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
         Left            =   1140
         Top             =   930
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1560
         Top             =   930
      End
      Begin VB.Timer tmrConn 
         Left            =   660
         Top             =   930
      End
      Begin VB.Frame fraVision 
         Caption         =   " VISION "
         Height          =   735
         Left            =   2970
         TabIndex        =   47
         Top             =   660
         Width           =   1965
         Begin VB.TextBox txtLastSeq 
            Appearance      =   0  '���
            Height          =   315
            Left            =   210
            TabIndex        =   48
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.TextBox txtSeqNo 
         Height          =   270
         Left            =   4410
         TabIndex        =   46
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtPosNo 
         Height          =   270
         Left            =   3720
         TabIndex        =   45
         Top             =   180
         Width           =   645
      End
      Begin VB.TextBox txtRackNo 
         Height          =   270
         Left            =   3000
         TabIndex        =   44
         Top             =   180
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
         Left            =   750
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   2040
         Top             =   930
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
               Picture         =   "frmInterface.frx":424A
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":47E4
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":4D7E
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":5318
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":5BAA
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":5D04
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":5E5E
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":5FB8
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":6892
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSCotrol.CButton cmdMatch 
         Height          =   405
         Left            =   3150
         TabIndex        =   71
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BackColor       =   12553049
         Caption         =   "��ġ���"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
      Begin VB.Image imgOn 
         Height          =   480
         Left            =   270
         Picture         =   "frmInterface.frx":716C
         Top             =   1500
         Width           =   480
      End
      Begin VB.Image imgOff 
         Height          =   480
         Left            =   150
         Picture         =   "frmInterface.frx":7A36
         Top             =   1440
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   6030
      TabIndex        =   25
      Top             =   720
      Width           =   16335
      _Version        =   393216
      _ExtentX        =   28813
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16645631
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmInterface.frx":8300
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6495
      Left            =   14730
      TabIndex        =   27
      Top             =   1650
      Width           =   6495
      _Version        =   393216
      _ExtentX        =   11456
      _ExtentY        =   11456
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      MaxCols         =   13
      MaxRows         =   20
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmInterface.frx":A214
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   30
      Left            =   10290
      TabIndex        =   63
      Top             =   2340
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      Format          =   51970049
      CurrentDate     =   44029
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '����
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   22260
      TabIndex        =   29
      Top             =   0
      Width           =   22260
   End
   Begin VB.PictureBox picTop 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H00BF8B59&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   22260
      TabIndex        =   28
      Top             =   30
      Width           =   22260
      Begin HSCotrol.CButton cmdBarcode 
         Height          =   495
         Left            =   9060
         TabIndex        =   56
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "��ü���/ã�� ��"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
      Begin HSCotrol.CButton cmdClear 
         Height          =   495
         Left            =   4680
         TabIndex        =   50
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "ȭ������"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":AE95
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin VB.Frame fraWorkList 
         Appearance      =   0  '���
         BackColor       =   &H00BF8B59&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   20760
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   3825
         Begin HSCotrol.CButton cmdWorkSave 
            Height          =   495
            Left            =   60
            TabIndex        =   60
            Top             =   90
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   873
            BackColor       =   12553049
            Caption         =   "����ȭ�� ��ũ����"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
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
            TabIndex        =   61
            Top             =   90
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   873
            BackColor       =   12553049
            Caption         =   "�����ũ �ҷ�����"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
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
         Begin VB.Shape Shape2 
            BackColor       =   &H00ACFFEF&
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   0  '����
            Height          =   645
            Left            =   0
            Top             =   0
            Width           =   3825
         End
      End
      Begin VB.Frame fraBarcode 
         Appearance      =   0  '���
         BackColor       =   &H00BF8B59&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   12510
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   6975
         Begin HSCotrol.CButton cmdBarFind 
            Height          =   495
            Left            =   5550
            TabIndex        =   59
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            BackColor       =   12553049
            Caption         =   "��üã��"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
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
            TabIndex        =   58
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            BackColor       =   12553049
            Caption         =   "��ü���"
            ForeColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
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
         Begin VB.CheckBox chkAdd 
            Appearance      =   0  '���
            BackColor       =   &H00AE8B59&
            Caption         =   "��ü��ȣ ����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   40
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtBarNum 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2190
            TabIndex        =   39
            Text            =   "123456789012345"
            Top             =   210
            Width           =   2175
         End
         Begin VB.TextBox txtOldBarNum 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   38
            Top             =   300
            Width           =   1665
         End
         Begin VB.Label lblRow 
            BackStyle       =   0  '����
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   42
            Top             =   90
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label10 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   ">>"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1830
            TabIndex        =   41
            Top             =   360
            Width           =   285
         End
      End
      Begin VB.Frame fraPatInfo 
         Appearance      =   0  '���
         BackColor       =   &H00BF8B59&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   30
         TabIndex        =   30
         Top             =   0
         Width           =   4635
         Begin VB.Shape Shape18 
            BackColor       =   &H00ACFFEF&
            BorderColor     =   &H00FFFFFF&
            Height          =   495
            Left            =   30
            Top             =   90
            Width           =   4575
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '����
            Caption         =   "��ü��ȣ :"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   36
            Top             =   150
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '����
            Caption         =   "�̸� :"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   35
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label14 
            BackStyle       =   0  '����
            Caption         =   "�˻����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   34
            Top             =   150
            Width           =   525
         End
         Begin VB.Label lblBarcode 
            BackStyle       =   0  '����
            Caption         =   "123456789012345"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   33
            Top             =   150
            Width           =   2805
         End
         Begin VB.Label lblPatNm 
            BackStyle       =   0  '����
            Caption         =   "ȫ��� (M/77)"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   32
            Top             =   360
            Width           =   2805
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  '����
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   31
            Top             =   180
            Width           =   1215
         End
      End
      Begin HSCotrol.CButton cmdSave 
         Height          =   495
         Left            =   7560
         TabIndex        =   51
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "��������"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":C117
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   495
         Left            =   1470
         TabIndex        =   52
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   " ������"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":C271
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton cmdDelete 
         Height          =   495
         Left            =   2940
         TabIndex        =   53
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   " �������"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":C3CB
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton CButton1 
         Height          =   495
         Left            =   240
         TabIndex        =   54
         Top             =   30
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "ȭ������"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":C525
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin HSCotrol.CButton cmdView 
         Height          =   495
         Left            =   6120
         TabIndex        =   55
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "�󼼰��"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":D7A7
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   495
         Left            =   10770
         TabIndex        =   57
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   873
         BackColor       =   12553049
         Caption         =   "��ũ����/�ε� ��"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
   Begin VB.Frame fraWorkInfo 
      Appearance      =   0  '���
      BackColor       =   &H00BF8B59&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   60
      TabIndex        =   18
      Top             =   720
      Width           =   5895
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   540
         TabIndex        =   19
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131792897
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131792897
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   345
         Left            =   3660
         TabIndex        =   62
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         BackColor       =   12553049
         Caption         =   "��ũ��ȸ"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         HoverColor      =   12648447
      End
      Begin HSCotrol.CButton cmdAll 
         Height          =   345
         Left            =   4770
         TabIndex        =   49
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         BackColor       =   12553049
         Caption         =   "�ϰ����"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         HoverColor      =   12648447
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "~"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   1980
         TabIndex        =   22
         Top             =   180
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸ�Ⱓ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   90
         Width           =   450
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  '�Ʒ� ����
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   22260
      TabIndex        =   0
      Top             =   9705
      Width           =   22260
      Begin HSCotrol.CButton cmdXML 
         Height          =   405
         Left            =   19140
         TabIndex        =   72
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         BackColor       =   15698777
         Caption         =   "XML ����"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
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
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   165
         Left            =   13620
         TabIndex        =   24
         Top             =   210
         Width           =   5325
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   13440
         Top             =   90
         Width           =   5685
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   4620
         Picture         =   "frmInterface.frx":E081
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   7455
         Picture         =   "frmInterface.frx":E60B
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   8730
         Picture         =   "frmInterface.frx":EB95
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��ſ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   5070
         TabIndex        =   5
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�޴½�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   6645
         TabIndex        =   4
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�����½�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   7770
         TabIndex        =   3
         Top             =   210
         Width           =   900
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmInterface.frx":F11F
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmInterface.frx":F269
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmInterface.frx":F3B3
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Com1 ���Ἲ��"
         ForeColor       =   &H00C0FFC0&
         Height          =   165
         Left            =   9180
         TabIndex        =   2
         Top             =   210
         Width           =   4125
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   4215
      End
      Begin VB.Label lblDBStatus 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�����ͺ��̽� ���Ἲ��"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   750
         TabIndex        =   1
         Top             =   180
         Width           =   3525
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   4410
         Top             =   90
         Width           =   9045
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7425
      Left            =   30
      TabIndex        =   26
      Top             =   1260
      Visible         =   0   'False
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   13097
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmInterface.frx":F4FD
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
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

Dim pDel                As Boolean
Dim strOldBarno         As String
Dim gMnuIdx             As Integer

Dim AckOn   As Boolean
Dim Sample_Seq  As String
Dim aMod    As String
Dim iIID    As String

'Private Sub cmdEnd_Click()
'
'    If MsgBox("���� ������Դϴ�. �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then
'
'        If comEqp.PortOpen = True Then
'            comEqp.PortOpen = False
'        End If
'
''        If gDBTYPE <> "99" Then
''            Call DisConnect_Server
''
''            Call DisConnect_Local
''        End If
'
'        Unload Me
'
'        'End
'    End If
'
'End Sub


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
                    spdOrder.MaxRows = spdOrder.MaxRows + 1
                    intRow = spdOrder.MaxRows
                    'For i = colCHECKBOX To colSTATE
                    For i = colEXAMDATE To colSTATE
                        Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                    Next

                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To spdOrder.MaxCols
                            spdOrder.Row = 0
                            spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(spdOrder.Text) Then
                                .Row = spdOrder.MaxRows
                                Call SetText(spdOrder, "��", spdOrder.MaxRows, intOCol)
                            End If
                        Next
                    Next

                    spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
        .MaxRows = 0
    End With

End Sub

Private Sub cmdBarcode_Click()

    If cmdBarcode.Caption = "��ü���/ã�� ��" Then
        fraBarcode.Visible = True
        cmdBarcode.Caption = "��ü���/ã�� ��"
        
        fraWorkList.LEFT = fraBarcode.LEFT + fraBarcode.WIDTH + 30
    
    Else
        fraBarcode.Visible = False
        cmdBarcode.Caption = "��ü���/ã�� ��"
    
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


'Private Sub cmdGetRslt_Click()
'    Dim strSendData As String
'    Dim strFirstSeq  As String
'    Dim strLastSeq  As String
''    Dim db_tmp As String * 100
'
'On Error GoTo RST
'
'    strFirstSeq = txtLastSeq.Text
'    strFirstSeq = (strFirstSeq - 1) - (txtRCnt.Text - 1)
'
'    strLastSeq = strFirstSeq + (txtRCnt.Text - 1)
'
'    strSendData = "0" & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq & vbLf
'
'    wSck.SendData strSendData
'    SetRawData "[Tx]" & strSendData
'
'Exit Sub
'
'RST:
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_cmdGetRslt_Click" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
'
'End Sub

Private Sub cmdHide_Click()
    
    spdResult.Visible = False
    'fraPatInfo.Visible = False
    
    spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH + 500

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
        MsgBox "��ũ����Ʈ���� �ϳ��� ��ü�� �����ϼ���", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    intCnt = 0
    
    For intORow = 1 To spdOrder.MaxRows
        If GetText(spdOrder, intORow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
            intOSrcRow = intORow
            blnSame = True
            'Exit For
        End If
    Next
    
    If blnSame = False Then
        MsgBox "�������Ʈ���� ��� ��ü�� �����ϼ���", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If intCnt > 1 Then
        MsgBox "�������Ʈ���� �ϳ��� ��ü�� �����ϼ���", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If blnSame = True Then
        For i = colCHECKBOX To colSTATE
            Call SetText(spdOrder, GetText(spdWork, intWSrcRow, i), intOSrcRow, i)
        Next
        
        varItems = GetText(spdWork, intWSrcRow, colITEMS)
        varItems = Split(varItems, "/")
        For intItems = 0 To UBound(varItems)
            For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                spdOrder.Row = 0
                spdOrder.Col = intOCol
                If varItems(intItems) = Trim(spdOrder.Text) Then
                    Call SetText(spdOrder, "��", intOSrcRow, intOCol)
                End If
            Next
        Next
        
        '��������
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
        SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, intOSrcRow, colEXAMTIME)) & "'  " & vbCrLf
        SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intOSrcRow, colSAVESEQ)) & vbCrLf
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- ����
        End If
    End If
End Sub

Private Sub cmdOrder_Click()
    Dim i As Integer
    
    strState = ""
    
    With spdOrder
        If .MaxRows > 0 Then
            For i = 1 To .MaxRows
                .Row = i
                .Col = colCHECKBOX
                If .Value = "1" And GetText(spdOrder, i, colSTATE) = "" Then
                    If MsgBox("�غ�� ������ �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo) = vbYes Then
                        Call SendData(ENQ)
                        strState = "Q"
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Next
        End If
    End With
    
End Sub

Private Sub cmdRcv_Click()
    
    pBuffer = txtRcv.Text

    Call ReceiveProcess

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

Private Sub cmdReceive_Click()
    Dim strInfoPath     As String
    Dim strRsltPath     As String
    Dim txtFilename     As String
    
    With CommonDialog1
        .CancelError = True
        
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly
        .InitDir = gComm.RSTPATH
        .Filter = "XML Files (info.xml)|*.xml|All Files (*.*)|*.*|"
        .FilterIndex = 1
        .Filename = ""
        .ShowOpen
        txtFilename = .Filename
    End With

    Screen.MousePointer = 11
    
    strInfoPath = txtFilename 'gComm.RSTPATH & "\" & "info.xml"
    
    Call DisplayNode_Info(strInfoPath)

    
    If UBound(strRecvData) > 1 Then
        Call SerialRcvData_MULTIPLATE
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
    
ErrHandler:
  ' ����ڰ� [���] ���߸� �������ϴ�.
Exit Sub
    
End Sub


Public Sub DisplayNode_Info(asPath As String)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    
    Dim i, j, k         As Integer
    Dim MsgType         As String
    
    On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    'xmlDoc.Load "D:\������Ʈ\VB\__JC�޵���\���򺴿�_MCC\IF\XML"
    
    k = 0
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                Exit For
            End If
        Next
        
        Erase strRecvData
        intBufCnt = 1
        ReDim Preserve strRecvData(7)
        strRecvData(intBufCnt) = "H|" & xNode.childNodes.Item(0).baseName
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "O|" & "1" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(4).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Runtime" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(8).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AreaUnderCurve" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(9).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Aggregation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(10).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "Velocity" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(11).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1
        strRecvData(intBufCnt) = strRecvData(intBufCnt) & "R|" & "AdditionalInformation" & "|" & xNode.childNodes.Item(3).childNodes.Item(0).childNodes.Item(12).childNodes.Item(1).nodeTypedValue
        intBufCnt = intBufCnt + 1


        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    MsgBox "���Ͽ���"
    
    Exit Sub
    
End Sub
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
'    'xmlDoc.Load "D:\������Ʈ\VB\���������������ǿ�\����\Result.xml"
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
'    MsgBox "���Ͽ���"
'    Exit Sub
'
'End Sub

Private Sub cmdResult_Click()

    frmResult.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("������ ����� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "�������") = vbYes Then
        With spdOrder
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = colCHECKBOX
                If .Value = 1 Then
                    Res = SaveTransData(lRow, spdOrder)
                    
                    If Res = -1 Then
                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdOrder, "�������", lRow, colSTATE
                    
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                    
                    Else
                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdOrder, "����Ϸ�", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                        
                    End If
                    spdOrder.Row = lRow
                    spdOrder.Col = colCHECKBOX
                    spdOrder.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub


Private Sub cmdSearch_Click()
    Dim i       As Integer
    Dim intRackNo   As Integer
    Dim intPosNo    As Integer
    Dim intSeq      As Integer
        
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork, Format(txtFrNo.Text, "0000"), Format(txtToNo.Text, "0000"))

End Sub

Private Sub cmdSend_Click()
    
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub


Private Sub cmdView_Click()
    
    If gWORKPOS = "M" Then
        If spdResult.Visible = False Then
            spdResult.Visible = True
            
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
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
            
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
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

Private Sub cmdWork_Click()

    frmWorkList.Show vbModal
    
End Sub

Private Sub cmdWorkList_Click()

    If cmdWorkList.Caption = "��ũ����/�ε� ��" Then
        fraWorkList.Visible = True
        cmdWorkList.Caption = "��ũ����/�ε� ��"
    Else
        fraWorkList.Visible = False
        cmdWorkList.Caption = "��ũ����/�ε� ��"
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
        If MsgBox("���� ȭ���� ����� ��ũ����Ʈ�� �ҷ����ڽ��ϱ�?", vbYesNo + vbInformation, "��ũ����Ʈ �ҷ�����") = vbNo Then
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
    'Ŭ������ ī��
    spdOrder.ClipboardCopy
    
    strBuffer = Clipboard.GetText()
    
    Call SetWorkData(strBuffer, spdOrder.MaxRows)

End Sub

Private Sub ReceiveProcess()
    
                    
    '>> RS232C
    If gComm.COMTYPE = "1" Then
        
        Select Case UCase(gHOSP.MACHNM)
            Case "INDIKO":          Call Phase_Serial_INDIKO
            
            Case "MINIVIDAS":       Call Phase_Serial_MINIVIDAS
            Case "THUNDERBOLT":     Call Phase_Serial_THUNDERBOLT
            Case "XN1000":          Call Phase_Serial_XN1000
            Case "MEDONIC":         Call Phase_Serial_MEDONIC
            Case "UROMETER120":     Call Phase_Serial_UROMETER720
            Case "HORIBA":          Call Phase_Serial_HORIBA
            Case "RP500":           Call Phase_Serial_RP500
            Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
            Case "URINSCAN":        Call Phase_Serial_URINSCAN
            Case "AVL9180":         Call Phase_Serial_AVL9180
            Case "CA800_ASTM":      Call Phase_Serial_CA800_ASTM
            Case "CA800":           Call Phase_Serial_CA800
            Case "AU480":           Call Phase_Serial_AU480
            Case "ACCESS2":         Call Phase_Serial_ACCESS2
            Case "HITACHI7020":     Call Phase_Serial_HITACHI7020
            Case "YUMIZEN":         Call Phase_Serial_YUMIZEN           '���ΰ��� HORIBA YUMIZEN H500
            Case "XP300":           Call Phase_Serial_XP300
            Case "ISMART30":        Call Phase_Serial_ISMART30
            Case "STAGO":           Call Phase_Serial_STAGO
            Case "PATHFAST":        Call Phase_Serial_PATHFAST
            Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
            'Case "KLITE":           Call Phase_Serial_KLITE
        End Select
        
    '>> SOCKET
    ElseIf gComm.COMTYPE = "2" Then
        Select Case UCase(gHOSP.MACHNM)
            Case "INDIKO":          Call Phase_TCP_INDIKO
            
            Case "GENEXPERT":       Call Phase_TCP_GENEXPERT
            Case "PPC300N":         Call Phase_TCP_PPC300N
            Case "KLITE":           Call Phase_TCP_KLITE
            Case "XP300":           Call Phase_TCP_XP300
            Case "YUMIZEN":         Call Phase_TCP_YUMIZEN
            Case "VISION":          Call Phase_TCP_VISION
            
        End Select
    End If
        
End Sub

Private Sub cmdXML_Click()
    
    Dim FindFile As String

    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
    If FindFile <> "" Then
        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"
        MsgBox "XML ���������� ���� �Ǿ����ϴ�.", vbOKOnly + vbInformation, Me.Caption
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
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
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
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If

        Case comEvCTS
            EVMsg$ = "CTS ���� ����"
        Case comEvDSR
            EVMsg$ = "DSR ���� ����"
        Case comEvCD
            EVMsg$ = "CD ���� ����"
        Case comEvRing
            EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF
            EVMsg$ = "EOF ����"

        '���� �޽���
        Case comBreak
            ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO
            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB
            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO
            ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame
            ERMsg$ = "�����̹� ����"
        Case comOverrun
            ERMsg$ = "�и�Ƽ ����"
        Case comRxOver
            ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity
            ERMsg$ = "�и�Ƽ ����"
        Case comTxFull
            ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else
            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select

    If ERMsg$ <> "" Then
        lblIFStatus.Caption = ERMsg$
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
'    If MsgBox("���� ������Դϴ�. �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then
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



Private Sub GetOrder_HITACHI7180(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        
        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        
        mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = strItems
        
            'GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
            
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
            
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            'GetOrder = STX & ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX
        
            strSndMsg = ";" & mOrder.Func & " 88" & Mid(mOrder.Order, 1, 88) & "100000" & LEFT(mOrder.PID & Space(30), 30)
            
            GetOrder = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If


        Call SendData(GetOrder)
        
        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_HITACHI7020(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarNo, intRow)
        
        '���ڵ带 ������� ���� ��쿡 ����Ѵ�.
        If gHOSP.BARUSE <> "Y" Then
            mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        End If
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
            
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 37" & Mid(mOrder.Order, 1, 37) & "00000" & ETX
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_STAGO(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        strItems = GetEquipExamCode_STAGO(gHOSP.MACHCD, pBarNo, intRow)
        
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If

        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_XN1000(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_XN1000(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If


        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_THUNDERBOLT(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_THUNDERBOLT(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- üũ�ڽ� ǥ��
            Call SetText(spdOrder, "0", intRow, colCHECKBOX)
            
            '-- �������(Order) ǥ��
            Call SetText(spdOrder, "��������", intRow, colSTATE)
        
            '-- ���� ������ ����
            Call SetText(spdOrder, "", intRow, colSPECIMEN)
        
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- üũ�ڽ� ǥ��
            Call SetText(spdOrder, "1", intRow, colCHECKBOX)
            
            '-- �������(Order) ǥ��
            Call SetText(spdOrder, "�����غ�", intRow, colSTATE)
            
            '-- ���� ������ ����
            Call SetText(spdOrder, strItems, intRow, colSPECIMEN)
            
        End If


        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_INDIKO(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        'Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_INDIKO(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            '-- �������(Order) ǥ��
            Call SetText(spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
            '-- �������(Order) ǥ��
            Call SetText(spdOrder, "�����غ�", intRow, colSTATE)
        End If

        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_CA800(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_CA800(gHOSP.MACHCD, pBarNo, intRow)
        
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
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            SendBuf = SendBuf & Space(15)
            SendBuf = SendBuf & "C"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            SendBuf = SendBuf & Right(Space(15) & mOrder.BarNo, 15)
            SendBuf = SendBuf & "B"
            SendBuf = SendBuf & Space(11)
            SendBuf = SendBuf & strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If

        SendBuf = STX & SendBuf & ETX
        
        Call Sleep(500)
        
        Call SendData(SendBuf)

        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_ACCESS2(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, pBarNo, intRow)
        
        mOrder.Order = strItems
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If

        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder_YUMIZEN(ByVal pBarNo As String, ByVal pType As String)

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
    
    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        strItems = GetEquipExamCode_YUMIZEN(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            'mOrder.Order = ""
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            'mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
        End If

        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub GetOrder(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String

    intRow = -1

    '-- 1. �������� ��ȸ
    With frmInterface
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmInterface.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmInterface.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    
        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, spdOrder)

        .spdOrder.RowHeight(-1) = 15

        strItems = ""
        mOrder.Order = ""
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarNo, intRow)

        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Or mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
            
        Else
            mOrder.NoOrder = False
        
            '-- �������(Order) ǥ��
            Call SetText(frmInterface.spdOrder, "��������", intRow, colSTATE)
            strOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & strItems & ETX
        End If

        Call SendData(strOrder)
        
        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub SendData(ByVal pSendData As Variant)

    '-- ����
    comEqp.Output = pSendData
    
    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    DoEvents
    
    '-- �αױ��
    Call SetRawData("[TxB]" & pSendData & "[TxE]")

End Sub

Private Sub SendWSckData(ByVal pSendData As Variant)

    '-- ����
    wSck.SendData pSendData
    
    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    DoEvents
    
    '-- �αױ��
    Call SetRawData("[Tx]" & pSendData)

End Sub


Private Sub TCPRcvData_KLITE()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
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
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- �������̽� ����
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
                    
                    '-- �˻縶���� ���� ��������
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
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
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
                        
                            '--- �������
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
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "�����", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
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
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'Public Sub TCPRcvData_BS200()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '������ Data
'    Dim strType         As String   '������ Record Type
'    Dim strBarno        As String   '������ ���ڵ��ȣ
'    Dim strSeq          As String   '������ Sequence
'    Dim strRackNo       As String   '������ Rack Or Disk No
'    Dim strTubePos      As String   '������ Tube Position
'    Dim strIntBase      As String   '������ ������ �˻��
'    Dim strMachResult   As String   '������ �����
'    Dim strResult       As String   '������ ���(����)
'    Dim strIntResult    As String   '������ ���(����)
'    Dim strQCResult     As String   '������ ���(QC)
'    Dim varResult       As Variant
'    Dim strFlag         As String   '������ Abnormal Flag
'    Dim strComm         As String   '������ Comment
'    Dim intCnt          As Integer
'
'    Dim strOrderCode    As String   'ó���ڵ�
'    Dim strTestCode     As String   '�˻��ڵ�
'    Dim strTestSubCode  As String   '�˻��ڵ�
'    Dim strTestName     As String   '�˻��
'    Dim strSeqNo        As String   '����DB �˻�Seq
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
'    Dim strRstRow       As String   '����������� ���� Row
'    Dim strDecYN        As String   '�����������
'    Dim strJudge        As String   '�������
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
'    '-- LDL ����
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
'                    'Resut Type(16)         : '' (����), 0 (Sample) , 1 (Calib. Result)
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
'                        '-- �˻������� ACK
'                        Case "ORU^R01"  '==> ACK^R01
'                                           strSndBuffer = SB & "MSH|^~\&|||" & mOrder.BSMaker & "|" & mOrder.BSMchNm & "|" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|" & mOrder.MSHMessageID & "|" & mOrder.MSHProduct & "|" & mOrder.MSHHL7Version & "||||0||ASCII|||" & vbCr
'                            strSndBuffer = strSndBuffer & "MSA|AA|" & mOrder.MSHMessageID & "|Message accepted|||0|" & vbCr
'                            strSndBuffer = strSndBuffer & EB & vbCr
'
'                            SetRawData "[Tx]" & strSndBuffer
'                            wSck.SendData strSndBuffer
'                        '-- ������û����
'                        Case "QRY^Q02"  '==> QCK^Q02
'
'                            strSndBuffer = ""
'
'                            With spdOrder
'                                For i = 1 To .MaxRows
'                                    If Trim(GetText(spdOrder, i, colCHECKBOX)) = "1" And Trim(GetText(spdOrder, i, colSTATE)) = "" Then
'                                        '-- ��������
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
'                            '-- ��������
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
'                        '-- ���� ����
'                        Case "ACK^Q03"
'                            '-- ������������
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
'                    '-- ���ʿ�������
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
'                    'Sample���
'                    If mOrder.MSHResultType = "0" Then
'                        strSeq = Trim$(mGetP(strRcvBuf, 4, "|"))
'
'                        If strBarno = "" Then
'                            strBarno = strSeq
'                        End If
'                    Else
'                        'cal ����� ó������
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
'                    '-- CREA �������
'                    If Trim(strIntBase) = "CRE" Then
'                        strGFR = ""
'                        strResult = Format(strResult, "##0.00")
'                        strCREA = strResult
'
'                        If CCur(strResult) > 0 Then
'                            '18�� �̻� ����
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
'                        'ó���� �������
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
'                            '-- ���Row �߰�
'                            strRstRow = .spdResult.DataRowCnt + 1
'                            If .spdResult.MaxRows < strRstRow Then
'                                .spdResult.MaxRows = strRstRow
'                            End If
'
'                            '�Ҽ��� ó��, ��� ���� ó��
'                            strMachResult = strResult
'                            If strQCTemp = "1" Then
'                                strResult = SetResult(strResult, strIntBase)
'                            End If
'                            strJudge = SetJudge(strResult, strIntBase)
'
'                            '������� ǥ��("���")
'                            SetText .spdOrder, "���", gRow, colSTATE
'
'                            '����� ǥ��
'                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
'                                    SetText .spdOrder, strResult, gRow, intCol
'
'                                    '�����ڵ�
'                                    strTestSubCode = gArrEQP(intCol - colSTATE, 17)
'
'                                    Exit For
'                                End If
'                            Next
'
'                            '-- ��� List
'                            SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '����
'                            SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          'ó���ڵ�
'                            SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '�˻��ڵ�
'                            SetText .spdResult, strTestSubCode, strRstRow, colRSUBCD          '�˻�SUB�ڵ�
'                            SetText .spdResult, strTestName, strRstRow, colRTESTNM            '�˻��
'                            SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '���ä��
'                            SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '�����
'                            SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS���
'                            SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '����
'                            SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '����ġ
'
'                            '-- ���� ����
'                            SetLocalDB gRow, strRstRow, "1", ""
'
'                            'strState = "R"
'
'                            '-- ���Count
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
'                    .spdResult.RowHeight(-1) = 14
'
'                    '-- GFR ����
'                    If strGFR <> "" Then
'                        strIntBase = "eGFR"
'                        strResult = strGFR
'                        strGFR = ""
'                        GoTo RST
'                    End If
'
'                    If strHbA1c <> "" And strIFCC <> "" And streAG <> "" Then
'                        '�Ұ������϶�
'                        strTotA1C = ""
'                        'strTotA1C = strTotA1C & "A1C : " & strA1C & vbCrLf
'                        strTotA1C = strTotA1C & "HbA1c% : " & strHbA1c & vbCrLf
'                        strTotA1C = strTotA1C & "IFCC : " & strIFCC & vbCrLf
'                        'strTotA1C = strTotA1C & "ADAG : " & strADAG & vbCrLf
'                        strTotA1C = strTotA1C & "eAG : " & streAG & vbCrLf
'                        strTotA1C = Mid(strTotA1C, 1, 254)
'
'                        'HbA1C����� �����Ѵ�.
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
'    'OBX|1|NM|HB|Hemoglobin|175.110862|��mol/L|-|N|||F||175.110862|20190829101649|||0||
'    'OBX|2|NM|A1C|Hemoglobin A1c|6.920507|��mol/L|-|N|||F||6.920507|20190829101649|||0||
'    'OBX|3|NM||HbA1c%|5.755654||-|N|||F||5.755654||||||
'    'OBX|4|NM||IFCC|39.409297||-|N|||F||39.409297||||||
'    'OBX|5|NM||ADAG|6.561490||-|N|||F||6.561490||||||
'    'OBX|6|NM||eAG|118.487267|mg/dL|-|N|||F||118.487267||||||
'
'
'
'        '## DB�� �������
'        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'            Res = SaveTransData(gRow, spdOrder)
'
'            If Res = -1 Then
'                '-- ���� ����
'                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                SetText .spdOrder, "�������", gRow, colSTATE
'            Else
'                '-- ���� ����
'                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
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
'                    '-- ����
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
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
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
'''                        '-- �������
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
'''                    '-- ���ȯ������
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
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                        
                    strOldBarno = strBarno
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strIntBase = mGetP(strRcvBuf, 3, "|")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = "" 'mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    
                    Call SetSQLData("RCV", strIntBase & ":" & strResult, "A")
                    
                    '-- MTB Ct�� ã��
'''                    If strIntBase = "^MTB-RIF^^MTB^^^Probe E^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 38 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "�����Ұ�"
'''                        End If
'''                    End If
'''
'''                    '-- TOX Ct�� ã��
'''                    If strIntBase = "^G3^^Toxi^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 5 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "�����Ұ�"
'''                        End If
'''                    End If
'''
'''                    '-- Carba-R �� ã��
'''                    If strIntBase = "^Carba-R^^IMP1^^^SPC^Ct" Then
'''                        strIntResult = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
'''                        If IsNumeric(strIntResult) Then
'''                            If strIntResult > 3 And strIntResult < 40 Then
'''                                strResult = "PASS"
'''                            Else
'''                                strResult = "FAIL"
'''                            End If
'''                        Else
'''                            strResult = "�����Ұ�"
'''                        End If
'''                    End If
                    
                    '-- �˻縶���� ���� ��������
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

                    .spdResult.RowHeight(-1) = 15
                
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
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ��� ������� �ʾ����� ������ �ǽɵǸ� Ÿ�˻� ����� Ȯ���Ͻñ� �ٶ��ϴ�."
                    
                        mResult.MTBRIF_CMTCD = "TB2"
                    
                    ElseIf strMTB = "DETECTED VERY LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ��� ����Ǿ� ������ ����ü �Ű����Դϴ�." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ� ���� �� ���ٱ� �󵵰� ������������ ����˴ϴ�." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB1"
                        
                    ElseIf strMTB = "DETECTED LOW" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ��� ����Ǿ� ������ ����ü �Ű����Դϴ�." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ� ���� �� ���ٱ� �󵵰� ������������ ����˴ϴ�." & vbNewLine

                        mResult.MTBRIF_CMTCD = "TB3"
                    
                    ElseIf strMTB = "DETECTED MEDIUM" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ��� ����Ǿ� ������ ����ü �Ű����Դϴ�." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ� ���� �� ���ٱ� �󵵰� ������������ ����˴ϴ�." & vbNewLine
                    
                        mResult.MTBRIF_CMTCD = "TB4"
                    
                    ElseIf strMTB = "DETECTED HIGH" Then
                        strMTBRIFCMT = ""
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ��� ����Ǿ� ������ ����ü �Ű����Դϴ�." & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "" & vbNewLine
                        strMTBRIFCMT = strMTBRIFCMT & "���ٱ� ���� �� ���ٱ� �󵵰� ������������ ����˴ϴ�." & vbNewLine
                        
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
                        strMTBRIFCMT = strMTBRIFCMT & "Rifamin �������� �������̰� ����Ǿ� �������� �Ǵܵ˴ϴ�." & vbNewLine
                    
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
                        strCarbaRCMT = strCarbaRCMT & "����� Carbapenemase �������� : strCarbaRPos" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "ȯ���� ��ü���� Carbapenemase �����ڰ� ����Ǿ����ϴ�." & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "Carbapenemase-producing Enter obacteriaceae (CPE) �����ڷ� �Ǵܵ˴ϴ�." & vbNewLine
                        
                    Else
                        strCarbaRCMT = ""
                        strCarbaRCMT = strCarbaRCMT & "[Comment]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "����� Carbapenemase �������� : ����" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "[Interpretation]" & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "�� �˻�� KPC, NDM, VIM �� OXA-48 �̿��� �˻翡�� carbapenemase�� ���ؼ� �߻��� CRE��," & vbNewLine
                        strCarbaRCMT = strCarbaRCMT & "�ʿ� �� CRE �������˻�(�˻��ڵ� : 40920)�� �Ƿ��Ͻñ� �ٶ��ϴ�." & vbNewLine
                        
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
        
        
        '## DB�� �������
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_TCPRcvData_GENEXPERT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_PPC300N()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
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
                    '-- �������̽� ����
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
                        '-- �������
                        With mResult
                            '.BarNo = strBarno
                            .Seq = strSeq
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ���ȯ������
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
                    
                    '-- �˻縶���� ���� ��������
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


                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_PPC300N" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub TCPRcvData_PPC300N_OLD()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strTypeSeq      As String   '������ Record Type Seq
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
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
                        '������û REQ;1
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
                                    Call SetText(spdOrder, "��������", i, colSTATE)
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
                        '�Ϲݻ��� REQ;2
                        '�������� REQ;3
                        'QC  ���� REQ;4
                        'Cal ���� REQ;5
                        
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
                            '-- �������
                            With mResult
                                .BarNo = strBarno
                                .RsltDate = Format(Now, "yyyy-mm-dd")
                                .RsltTime = Format(Now, "hh:mm:ss")
                                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                            End With
                        End If
                        
                        '-- ���ȯ������
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
                    
                    '��������
                    If strTypeSeq = "1" Then
                        '�������� TRA;1
                        strSeq = mGetP(strTemp, 1, ";")
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASK;2||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '2.SEND : <SB>|;^\|LisDemo|ASK;2||ASCII|<EB>
                        
                        strSend = SB & "|;^\|LisDemo|" & "ASW;" & strSeq & "||ASCII|" & EB
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                        '3.SEND : <SB>||;^\LisDemo|ASW;6|1234;ALT|ASCII|<EB>
                    
                    '�Ϲݰ������
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
                        
                        '-- �˻縶���� ���� ��������
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
                    'QC ����
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
                        
                    'Cal ����
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
                    
                    .spdResult.RowHeight(-1) = 15
                
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
            
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
            
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_F200()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
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
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- �������̽� ����
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
                    
                    '-- �˻縶���� ���� ��������
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
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
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
                        
                            '--- �������
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
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
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
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HITACHI7180()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
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
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY ����
                    
                    '-- ���� ����
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ���� ����
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
                    
                    '## Control, Calibration �����ʹ� ������
'                    If UCase(strFunc) = "K" Or UCase(strFunc) = "L" Or UCase(strFunc) = "G" Or UCase(strFunc) = "H" Then
'                        '-- ���� ����
'                        Call SendData(SndMore)
'                        strState = ""
'                        Exit Sub
'                    End If
                    
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
'''
'''                    If UCase(strFunc) = "F" Then
'''                        '-- ���� ����
'''                        Call SendData(SndMore)
'''                        strState = ""
'''                        Exit Sub
'''                    End If
                    
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
            
                        '-- CREA �������
                        If Trim(strIntBase) = "2" Then
                            strGFR = ""
                            strResult = Format(strResult, "##0.00")
                            strCREA = strResult
                            
                            If mPatient.AGE <> "" And mPatient.SEX <> "" Then
                                If CCur(strResult) > 0 Then
                                    '18�� �̻� ����
                                    If IsNumeric(strCREA) And mPatient.AGE > 18 Then
                                        If mPatient.SEX = "M" Then
                                            strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203)
                                        ElseIf mPatient.SEX = "F" Then
                                            strGFR = 186 * (strCREA ^ -1.154) * (mPatient.AGE ^ -0.203) * 0.742
                                        End If
                                        
                                        If strGFR <> "" Then
                                            strGFR = Format(strGFR, "##0.00")
                                            If strGFR <= 120 Then
                                                strGFR = Round(strGFR, 2)
                                            ElseIf strGFR > 120 Then
                                                strGFR = "> 120"
                                            End If
                                        End If
                                    End If
                                Else
                                    strGFR = "Error"
                                End If
                            End If
                        End If
                        
                        If strIntBase = "20" Then    'CRP
                            strCRP = strResult
                        End If
                        
                        If strIntBase = "21" Then    'RF
                            strRF = strResult
                        End If
                    
                        If strIntBase = "19" Then    'TCHO
                            strTC = strResult
                        End If

                        If strIntBase = "10" Then   'TG
                            strTG = strResult
                        End If

                        If strIntBase = "9" Then    'HDLC
                            strHDL = strResult
                        End If
                    
'                        If strIntBase = "25" Then    'GA
'                            strGA = strResult
'                        End If
'
'                        If strIntBase = "26" Then    'GA-Alb
'                            strGAAlb = strResult
'                        End If
                    
ReCal:
                        '-- �˻���ó�� ���μ���
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
                    '��񿡼� �������
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
'                        '-- �˻���ó�� ���μ���
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
                    
                    'LDL ���
                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                        strIntBase = "95"
                        strResult = strTC - ((strTG / 5) + strHDL)
                        If strResult < 0 Then
                            strResult = "0"
                        End If
                        strIntResult = ""
                        strTC = ""
                        strTG = ""
                        strHDL = ""
                        
                        '-- �˻���ó�� ���μ���
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
                    
                    'CRP ����
                    If strCRP <> "" Then
                        strIntBase = "87"
                        If strCRP < 0.5 Then
                            strResult = "Negative (" & strCRP & ")"
                        Else
                            strResult = "Positive (" & strCRP & ")"
                        End If
                        strIntResult = ""
                        strCRP = ""
                        
                        '-- �˻���ó�� ���μ���
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
                    
                    'RA Fact ����
                    If strRF <> "" Then
                        strIntBase = "88"
                        If strRF < 15 Then
                            strResult = "Negative (" & strRF & ")"
                        Else
                            strResult = "Positive (" & strRF & ")"
                        End If
                        strIntResult = ""
                        strRF = ""
                        
                        '-- �˻���ó�� ���μ���
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
                    
                    '-- GFR ����
                    If strGFR <> "" Then
                        strIntBase = "89"
                        strResult = strGFR
                        strIntResult = ""
                        strGFR = ""
                        
                        '-- �˻���ó�� ���μ���
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
    

                    
                    Call SendData(SndMore)
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_H7180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    Dim R As Integer
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
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    Dim R   As Integer
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
        
            '������ȣ, SeqNo
            strBarno = Trim(iPID)
            strSeqNo = Trim(rSeq)
            
            If Trim(strBarno) = "" Then Exit Sub
            
            '-- �������
            With mResult
                .BarNo = strBarno
                .Seq = strSeqNo
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
            
            '-- ���ȯ������
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++�� ��� ���˻��ڵ尡 �����ϱ� ������ Measured & Calibrated �� ������ �ʿ�...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
        
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                
                strResult = Mid(MsgBuf, x2, x1 - x2)
                strIntResult = strResult
                
                '-- �˻���ó�� ���μ���
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
                'Ca++�� ��� ���˻��ڵ尡 �����ϱ� ������ Measured & Calibrated �� ������ �ʿ�...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
    
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
                strIntResult = strResult
            
                '-- �˻���ó�� ���μ���
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
            
            .spdResult.RowHeight(-1) = 15
    
            '## DB�� �������
            If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                Res = SaveTransData(gRow, spdOrder)
    
                If Res = -1 Then
                    '-- ���� ����
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "�������", gRow, colSTATE
                Else
                    '-- ���� ����
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
    
                          SQL = "Update PATRESULT Set                                                               " & vbCrLf
                    SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                    SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                    SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                    SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                    SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- ����
                    End If
                End If
                strState = ""
            End If
        End If
    End With

Exit Sub

Err:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_RP500" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_HITACHI7020()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
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

            Select Case strType
                Case ">", "?", "@"      'ANY ����
                    
                    '-- ���� ����
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ���� ����
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
                    strFunc = Mid(strRcvBuf, 2, 1)              ' Function

                    If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                         Exit Sub
                    End If

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    If gHOSP.BARUSE = "Y" Then
                        '���ڵ� ���
                        sFunc = Mid(strRcvBuf, 2, 40)
                    Else
                        '���ڵ� �̻��
                        sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    End If
                    
                    With mOrder
                        .BarNo = strBarno
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                    End With
                    
                    Call GetOrder_HITACHI7020(Trim$(strBarno), gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    
                Case ":"    '## End
                
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    
                    '## Control, Calibration �����ʹ� ������
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    '## QC
                    If UCase(strFunc) = "F" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "O"
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If
                    
                    For ii = 45 To Len(strRcvBuf) Step 10
                        If strIntBase = "18" Then Stop
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        'strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 5))
                        strIntResult = strResult
                        
                        '-- �˻縶���� ���� ��������
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
                        
                        'RA����
                        If strIntBase = "20" Then
                            'RA����
                            strIntBase = "99"
                            If IsNumeric(strResult) Then
                                If strResult > 15 Then
                                    strResult = "Positive"
                                Else
                                    strResult = "Negative"
                                End If
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
                    
                    Call SendData(SndMore)
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_H7020" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Function ResultProcess(ByVal pBarNo As String, ByVal pIntBase As String, ByVal pResult As String, ByVal pIntResult As String) As Boolean
    Dim RS_L            As ADODB.Recordset
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strSeqNo        As String   '�˻����
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim strCheck        As String   '�˻����üũ
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    'Dim strIntResult    As String   '������ ���(����)
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strRstType      As String
    Dim i               As Integer
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCol          As Integer  '����÷� ����
    
    Dim strIntResult    As String   '��ȯ�� ��ġ���
    Dim strChrResult    As String   '��ȯ�� ���ڰ��
    Dim strResult       As String   '�������
    
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
    
    SQL = ""
    SQL = SQL & "SELECT TESTNAME,ABBRNAME,EQPMASTER.SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC,RESTYPE   " & vbCrLf
    SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7               " & vbCrLf
    SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7              " & vbCrLf
    SQL = SQL & "     , AMRLimit8,  AMRLimit9,  AMRLimit10,  AMRLimit11,  AMRLimit12,  AMRLimit13,  AMRLimit14          " & vbCrLf
    SQL = SQL & "     , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14         " & vbCrLf
    SQL = SQL & "     , AMRINResult                                                                                     " & vbCrLf
    SQL = SQL & ", (SELECT TOP 1 TESTMASTER.TESTCODE "
    SQL = SQL & "     FROM TESTMASTER "
    SQL = SQL & "    WHERE TESTMASTER.RSLTCHANNEL = EQPMASTER.RSLTCHANNEL"
    If gPatOrdCd <> "" Then
        SQL = SQL & "      AND TESTMASTER.TESTCODE in (" & gPatOrdCd & ") " & vbCrLf
    End If
    SQL = SQL & "  ) AS TESTCODE "
    SQL = SQL & "  FROM EQPMASTER LEFT JOIN AMRMASTER                                                                   " & vbCrLf
    SQL = SQL & "   ON (EQPMASTER.RSLTCHANNEL = AMRMASTER.RSLTCHANNEL)                                                  " & vbCrLf
    SQL = SQL & " WHERE EQPMASTER.EQUIPCD     = '" & gHOSP.MACHCD & "'                                                  " & vbCrLf
    SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = '" & pIntBase & "'                                                      " & vbCrLf
    'If gPatOrdCd <> "" Then
    '    SQL = SQL & "   AND EQPMASTER.TESTCODE in (" & gPatOrdCd & ") "
    'End If
    
    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        strSeqNo = Trim(RS_L.Fields("SEQNO"))
        strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
        strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
        strAbbrName = Trim(RS_L.Fields("ABBRNAME")) & ""
        
        '-- �Ҽ�����ȯ ��뿩�ο� ��ȯ�ڸ���
        intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
        intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
        
        '-- ������ ����ġ�� ���Ͽ� �����Ѵ�.
        If mPatient.SEX = "M" Then
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        ElseIf mPatient.SEX = "F" Then
            strLow = Trim(RS_L.Fields("REFFLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
        Else
            '-- ��������ġ�� �⺻���� �Ѵ�
            strLow = Trim(RS_L.Fields("REFMLOW")) & ""
            strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
        End If
                
        '�����  (0:��ġ,1:����,2:��ġ/����)
        strResType = Trim(RS_L.Fields("RESTYPE")) & ""
        strIntResult = ""
        
        '-- �˻����� ��ġ���ϰ��
        If strResType = 0 Then
            '--- �ο쵥���ͷ� �������
            strJudge = ""
            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                If IsNumeric(pIntResult) Then
                    If CCur(pIntResult) > CCur(strLow) And CCur(pIntResult) < CCur(strHigh) Then
                        strJudge = ""
                    ElseIf CCur(strHigh) <= CCur(pIntResult) Then
                        strJudge = "H"
                    ElseIf CCur(strLow) >= CCur(pIntResult) Then
                        strJudge = "L"
                    End If
                End If
            End If
            
            '-- �Ҽ��� ó��
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
            
            '-- �ο쵥���ͷ� AMR ���� (��ġ��)
            If IsNumeric(pIntResult) Then
                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT1"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT2"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT3"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT4"))
                    End If
                End If
                If strIntResult = "" Then
                    strIntResult = pIntResult
                End If
            Else
                strIntResult = pIntResult
            End If
            
            If strIntResult <> "" Then
                strResult = strIntResult
            Else
                strResult = pIntResult
            End If
            
            If strResult = "" Then
                strResult = pResult
            End If
            
        '-- �˻����� �������ϰ��
        ElseIf strResType = 1 Then
            If pResult <> "" Then
                '-- AMR ���� (������ �ܹ�)
                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT5"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT6"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT7"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT8"))
                    End If
                End If
                
                '-- AMR ���� (������ �幮)
                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT9"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT10"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT11"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT12"))
                    End If
                End If
                
                If strChrResult = "" Then
                    strChrResult = pResult
                End If
            Else
                strChrResult = pResult
            End If
        
            If strChrResult <> "" Then
                strResult = strChrResult
            Else
                strResult = pResult
            End If
            
            If strResult = "" Then
                strResult = pIntResult
            End If
            
        '-- �˻����� ��ġ+�������ϰ��
        ElseIf strResType = 2 Then
            '-- �Ҽ��� ó��
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
                pIntResult = Format(pIntResult, strResType)
            End If
                
            '-- AMR ���� (��ġ��)
            If IsNumeric(pIntResult) Then
                If Trim(RS_L.Fields("AMRLIMIT1")) & "" <> "" Then
                    If CCur(pIntResult) < CCur(Trim(RS_L.Fields("AMRLIMIT1"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT1"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT2")) & "" <> "" Then
                    If CCur(pIntResult) <= CCur(Trim(RS_L.Fields("AMRLIMIT2"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT2"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT3")) & "" <> "" Then
                    If CCur(pIntResult) > CCur(Trim(RS_L.Fields("AMRLIMIT3"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT3"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT4")) & "" <> "" Then
                    If CCur(pIntResult) >= CCur(Trim(RS_L.Fields("AMRLIMIT4"))) Then
                        strIntResult = Trim(RS_L.Fields("AMRRESULT4"))
                    End If
                End If
            
                If strIntResult = "" Then
                    strIntResult = pIntResult
                End If
            Else
                strIntResult = pIntResult
            End If
                                    
            '-- AMR ���� (������)
            If pResult <> "" Then
                '-- AMR ���� (������ �ܹ�)
                If Trim(RS_L.Fields("AMRLIMIT5")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT5")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT5"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT6")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT6")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT6"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT7")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT7")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT7"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT8")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT8")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT8"))
                    End If
                End If
                
                '-- AMR ���� (������ �幮)
                If Trim(RS_L.Fields("AMRLIMIT9")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT9")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT9"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT10")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT10")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT10"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT11")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT11")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT11"))
                    End If
                End If
                If Trim(RS_L.Fields("AMRLIMIT12")) & "" <> "" Then
                    If pResult = Trim(RS_L.Fields("AMRLIMIT12")) Then
                        strChrResult = Trim(RS_L.Fields("AMRRESULT12"))
                    End If
                End If
                
                If strChrResult = "" Then
                    strChrResult = pResult
                End If
            
            Else
                strChrResult = pResult
            End If
        
            '��ġ��� ����
            '0:������, 1:����(����), 2:����(����)
            If strIntResult <> "" And strChrResult <> "" Then
                If Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                    strResult = strChrResult & "(" & strIntResult & ")"
                ElseIf Trim(RS_L.Fields("AMRINResult") & "") = "1" Then
                    strResult = strIntResult & "(" & strChrResult & ")"
                End If
            Else
                If strChrResult <> "" Then
                    strResult = strChrResult
                ElseIf strIntResult <> "" Then
                    strResult = strIntResult
                End If
            End If
        End If
        
        With frmInterface
            '-- ���Row �߰�
            intRstRow = .spdResult.DataRowCnt + 1
            If .spdResult.MaxRows < intRstRow Then
                .spdResult.MaxRows = intRstRow
            End If
    
            '-- ������� ǥ��("���")
            SetText .spdOrder, "�����", gRow, colSTATE
    
            '-- ����ȭ�� ����� ǥ��
            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                If strAbbrName = gArrEQP(intCol - colSTATE, 6) Then
                    SetText .spdOrder, strResult, gRow, intCol
                    
                    '-- H/L ����ǥ��
                    If strJudge = "H" Then
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbRed
                    ElseIf strJudge = "L" Then
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbBlue
                    Else
                        .spdOrder.Row = gRow
                        .spdOrder.Col = intCol
                        .spdOrder.ForeColor = vbBlack
                    End If
                    
                    Exit For
                End If
            Next
    
            '-- ��� List
            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
            SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '����
            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
            SetText .spdResult, pIntBase, intRstRow, colRCHANNEL              '���ä��
            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
            
            '-- ������� ��ȸ
            strPrevRslt = GetPrevResult(mResult.BarNo, pIntBase, strTestCode)
            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
            
            '-- H/L ����ǥ��
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
            
            '-- ���Count
            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                SetText .spdOrder, "1", gRow, colRCNT
            Else
                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
            End If
        End With
        
        '-- ���� ����
        Call SetLocalDB(gRow, intRstRow, "1", "")
        
        ResultProcess = True
    
    End If
    
End Function


'Private Sub SerialRcvData_MEDONIC()
'    '��� ���� ����
'    Dim strRcvBuf       As String   '������ Data
'    Dim strType         As String   '������ Record Type
'    Dim strBarno        As String   '������ ���ڵ��ȣ
'    Dim strSeq          As String   '������ Sequence
'    Dim strRackNo       As String   '������ Rack Or Disk No
'    Dim strTubePos      As String   '������ Tube Position
'    Dim strIntBase      As String   '������ ������ �˻��
'    Dim strResult       As String   '������ ���(����)
'    Dim strIntResult    As String   '������ ���(����)
'    Dim strQCResult     As String   '������ ���(QC)
'    Dim strFlag         As String   '������ Abnormal Flag
'    Dim strComm         As String   '������ Comment
'
'    '������ ����
'
'    Dim intCnt          As Integer  '��� Frame ����
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
'    '���� ����
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
'                    '-- �������
'                    With mResult
'                        .BarNo = strBarno
'                        .RsltDate = Format(Now, "yyyy-mm-dd")
'                        .RsltTime = Format(Now, "hh:mm:ss")
'                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'                    End With
'
'
'                    '-- ���ȯ������
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
'                        '-- �˻���ó�� ���μ���
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
'                        .spdResult.RowHeight(-1) = 15
'                    End If
'                End If
'
'                If InStr(strRcvBuf, "</smpresults>") > 0 Then
'                    strState = "L"
'                End If
'
'
'                '## DB�� �������
'                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "L" Then
'                    Res = SaveTransData(gRow, spdOrder)
'
'                    If Res = -1 Then
'                        '-- ���� ����
'                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                        SetText .spdOrder, "�������", gRow, colSTATE
'                    Else
'                        '-- ���� ����
'                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                        SetText .spdOrder, "����Ϸ�", gRow, colSTATE
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
'                            '-- ����
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
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_MEDONIC" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
'
'
'End Sub


Private Sub SerialRcvData_XN1000()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                
                '-- ���ȯ������
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
                    '## ������� ����
                    strResult = mGetP(strTemp2, 2, "^")
                Else
                    '## ������� ����
                    strIntResult = strTemp2
                End If
                
'                    If strIntBase = "WBC" And IsNumeric(strResult) Then
'                        strWBC = strResult * 1000
'                    End If
'
'                    If strIntBase = "NEUT%" And IsNumeric(strResult) Then
'                        strNeut = strResult / 100
'                    End If
                
                
                '-- �˻���ó�� ���μ���
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
                '-- ����׸� ����
                m = 0
                For i = colSTATE + 1 To spdOrder.MaxCols
                    spdOrder.Row = 0
                    spdOrder.Col = i
                    If spdOrder.FontBold = True Then
                        '����׸� ä��ã��
                        strCalChannel = GetChannel(spdOrder.Text)

                        strCalCulate = GetCalContents(strCalChannel, "")
                        varCalCulate = Split(strCalCulate, "%")

                        For j = 0 To UBound(varCalCulate)
                            For k = colSTATE + 1 To spdOrder.MaxCols
                                spdOrder.Row = 0
                                spdOrder.Col = k
                                If spdOrder.Text = varCalCulate(j) Then
                                    strCalNm(m) = varCalCulate(j)
                                    strCalCon(m) = GetText(spdOrder, gRow, k)
                                    
                                    strCalCulate = Replace(strCalCulate, strCalNm(m), strCalCon(m))
                                    m = m + 1
                                    Exit For
                                End If
                            Next
                        Next
                        
                        If m > 0 Then
                            strCalCulate = Replace(strCalCulate, "%", "")
                            strIntResult = mCalP(strCalCulate)
                            
'
                            strIntBase = strCalChannel
                            '-- �˻���ó�� ���μ���
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
                            
                            
                        End If
                    End If
                Next
                
                '## DB�� �������
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- ���� ����
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "�������", gRow, colSTATE
                    Else
                        '-- ���� ����
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_THUNDERBOLT()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                    
                        '-- ���ȯ������
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
                    
                    
                    '-- �˻���ó�� ���μ���
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
                
                Case "L"
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_THUNDERBOLT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_MINIVIDAS()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim intCnt          As Integer  '��� Frame ����
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
        
        '-- ���ä��
        strTmp = Trim(strRData(5))
        If Mid(strTmp, 2, 2) <> "rt" Then
            Exit Sub
        End If
        strIntBase = Trim(Mid(strTmp, 4))
        
        '-- �˻���(����)
        strTmp = Trim(strRData(9))
        If Mid(strTmp, 2, 2) <> "ql" Then
            Exit Sub
        End If
        strResult = Trim(Mid(strTmp, 4))
        
        '-- �˻���(����)
        strTmp = Trim(strRData(10))
        If Mid(strTmp, 2, 2) <> "qn" Then
            Exit Sub
        End If
        strIntResult = Trim(Mid(strTmp, 4))
                
        '������� Flag
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
                    
        '-- ���ȯ������
        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
        
        If gRow <= 0 Then
            Exit Sub
        End If
                
        '-- �˻���ó�� ���μ���
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
    
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_THUNDERBOLT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_MULTIPLATE()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    strIntBase = mGetP(strRcvBuf, 2, "|")
                    strResult = mGetP(strRcvBuf, 3, "|")
                    '��������
                    strResult = mGetP(strResult, 1, " ")
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_XN1000" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_INDIKO()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
On Error GoTo RST

'    ReDim Preserve strRData(UBound(strRecvData))
    
'    strRData = strRecvData
    
    With frmInterface
        For intCnt = 1 To UBound(strRecvData)
            
            'strRcvBuf = RcvBuffer
            strRcvBuf = strRecvData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^1700000104^^||1^Analyzer 1^5.0|||||||P||20170120174819
                    '��ġ
                    '2Q|1|^1^^||^^^ALL^||||||||O
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
                    'strRackNo = mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^")
                    'strTubePos = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    
                    With mOrder
                        .NoOrder = False
                        '.Seq = strBarno
                        .BarNo = strBarno
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                    End With
                    
                    Call GetOrder_INDIKO(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    strQState = "Q"
                    
                Case "P"
                    '2P|1|1700000104|||||||||||||||||||||||||||||||
                    '��ġ
                    '2P|1|o12292|||o12292||||||||||||||||||||||||||||
                Case "O"
                    '3O|1|1700000104^0.0^2^1||^^^Alb^0.0|R||||||X||||1|||||||||1|F
                    '3O|1|2008060039^0.0^1^1||^^^ALB^0.0|R||||||X||||1|||||||||1|F
                    '��ġ
                    '3O|1|1^0.0^2^1||^^^ALB^0.0|R||||||X||||1|||||||||1|F

                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    If Len(strBarno) = 10 Then
                        strBarno = "20" & strBarno
                    End If
                    If strOldBarno <> strBarno Then
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        
                        '-- ���ȯ������
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    '4R|1|^^^Alb^0.0|4.4|g/dl|0^0^|N||F\R||||20170119170203|Analyzer 1
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    
                    If InStr(strTemp1, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strTemp1, 2, "^")
                    Else
                        '## ������� ����
                        strIntResult = strTemp1
                    End If
                    
                    '-- �˻���ó�� ���μ���
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

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_INDIKO()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
            
            'strRcvBuf = RcvBuffer
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                    '2Q|1|^1700000104^^||1^Analyzer 1^5.0|||||||P||20170120174819
                    '��ġ
                    '2Q|1|^1^^||^^^ALL^||||||||O
                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
                    'strRackNo = mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^")
                    'strTubePos = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    
                    With mOrder
                        .NoOrder = False
                        '.Seq = strBarno
                        .BarNo = strBarno
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                    End With
                    
                    Call GetOrder_INDIKO(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                    strQState = "Q"
                    
                Case "P"
                    '2P|1|1700000104|||||||||||||||||||||||||||||||
                    '��ġ
                    '2P|1|o12292|||o12292||||||||||||||||||||||||||||
                Case "O"
                    '3O|1|1700000104^0.0^2^1||^^^Alb^0.0|R||||||X||||1|||||||||1|F
                    '3O|1|2008060039^0.0^1^1||^^^ALB^0.0|R||||||X||||1|||||||||1|F
                    '��ġ
                    '3O|1|1^0.0^2^1||^^^ALB^0.0|R||||||X||||1|||||||||1|F

                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
                    If Len(strBarno) = 10 Then
                        strBarno = "20" & strBarno
                    End If
                    If strOldBarno <> strBarno Then
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        
                        '-- ���ȯ������
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case "R"
                    '4R|1|^^^Alb^0.0|4.4|g/dl|0^0^|N||F\R||||20170119170203|Analyzer 1
                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    
                    If InStr(strTemp1, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strTemp1, 2, "^")
                    Else
                        '## ������� ����
                        strIntResult = strTemp1
                    End If
                    
                    '-- �˻���ó�� ���μ���
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

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_CA800()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                
                '-- ���ȯ������
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
                                '   - PT %���� 100�̻��̸� �ǹ̾��� ����� "100�̻�"���� �������
                                '     �ϴ°����� ����
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
                    
                    '-- �˻���ó�� ���μ���
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
                
                .spdResult.RowHeight(-1) = 15

                '## DB�� �������
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- ���� ����
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "�������", gRow, colSTATE
                    Else
                        '-- ���� ����
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ������ Result Flags�� ���� �󼼼��� ��ȸ
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
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                    
                    '-- ���ȯ������
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
                        '## ������� ����
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ������� ����
                        strResult = strTemp2
                        strIntResult = strTemp2
                    End If
                    
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_CA800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XP300()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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

                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
                        '## ������� ����
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ������� ����
                        strIntResult = strTemp2
                    End If
                        
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_XP300" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_URINSCAN()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim Pos             As Integer
        
On Error GoTo RST
    
    With frmInterface
        Pos = InStr(RcvBuffer, "ID_NO")
        If Pos > 0 Then
            RcvBuffer = Replace(RcvBuffer, vbLf, "")
            strRecvData = Split(RcvBuffer, vbCr)
            
            '-- SEQ ��ȣ ã��
            strRcvBuf = strRecvData(2)
            strRcvBuf = mGetP(strRcvBuf, 2, ":")
            strRcvBuf = mGetP(strRcvBuf, 1, "-")
            strSeq = Trim(strRcvBuf)
            
            strBarno = strSeq
            With mResult
                .BarNo = strBarno
                .TubePos = strSeq
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
                    
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            
            For intCnt = 4 To UBound(strRecvData)
                strRcvBuf = strRecvData(intCnt)
                
                strType = Trim(Mid$(strRcvBuf, 1, 3))
                strIntBase = strType
                strResult = ""

                Select Case strType
                    Case "p.H", "pH", "S.G", "SG", "COL" '## �Ҽ��� ���� 3�ڸ�
                        strResult = Trim$(Mid$(strRcvBuf, 4))
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    
                    Case Else
                        strResult = Trim$(Mid$(strRcvBuf, 4, 7))
                        'strResult = Trim(Mid(strRcvBuf, 12))  '-- ����
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                        
                End Select
                
                strIntResult = strResult
                        
                '-- �˻���ó�� ���μ���
                If strIntBase <> "" And strResult <> "" Then
                    If strState = "" Or strState = "O" Then
                        strState = ""
                    End If
                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strIntResult) = True Then
                        strState = "R"
                    Else
                        'ȭ��ǥ��
                        If strState = "" Then
                            strState = ""
                        End If
                        
                        Dim RS_L        As ADODB.Recordset
                        Dim strSeqNo    As String
                        Dim strTestCode As String
                        Dim strTestName As String
                        Dim strAbbrName As String
                        Dim intRstRow   As Integer
                        Dim intCol      As Integer
                        
                        SQL = ""
                        SQL = SQL & "SELECT EQPMASTER.TESTCODE,TESTNAME,ABBRNAME,EQPMASTER.SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC,RESTYPE   " & vbCrLf
                        SQL = SQL & "     , AMRLimit1,  AMRLimit2,  AMRLimit3,  AMRLimit4,  AMRLimit5,  AMRLimit6,  AMRLimit7               " & vbCrLf
                        SQL = SQL & "     , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7              " & vbCrLf
                        SQL = SQL & "     , AMRLimit8,  AMRLimit9,  AMRLimit10,  AMRLimit11,  AMRLimit12,  AMRLimit13,  AMRLimit14          " & vbCrLf
                        SQL = SQL & "     , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14         " & vbCrLf
                        SQL = SQL & "     , AMRINResult                                                                                     " & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER , AMRMASTER                                                                           " & vbCrLf
                        SQL = SQL & " WHERE EQPMASTER.EQUIPCD     = '" & gHOSP.MACHCD & "'                                                            " & vbCrLf
                        SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = '" & strIntBase & "'                                                                " & vbCrLf
                        SQL = SQL & "   AND EQPMASTER.EQUIPCD     = AMRMASTER.EQUIPCD                                                       " & vbCrLf
                        SQL = SQL & "   AND EQPMASTER.RSLTCHANNEL = AMRMASTER.RSLTCHANNEL                                                   " & vbCrLf
                        SQL = SQL & "   AND EQPMASTER.TESTCODE    = AMRMASTER.TESTCODE                                                      "
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqNo = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            strAbbrName = Trim(RS_L.Fields("ABBRNAME")) & ""
                        
                            With frmInterface
                                '-- ���Row �߰�
                                intRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < intRstRow Then
                                    .spdResult.MaxRows = intRstRow
                                End If
                        
                                '-- ������� ǥ��("���")
                                SetText .spdOrder, "�����", gRow, colSTATE
                        
                                '-- ����ȭ�� ����� ǥ��
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strAbbrName = gArrEQP(intCol - colSTATE, 6) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        Exit For
                                    End If
                                Next
                        
                                '-- ��� List
                                SetText .spdResult, strSeqNo, intRstRow, colRSEQNO                  '����
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                                SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                                SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                                SetText .spdResult, strResult, intRstRow, colRMACHRESULT        '�����
                                SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                                
                            End With
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
                                                    
                        End If
                    End If
                End If
                
                .spdResult.RowHeight(-1) = 15
            Next
                
            '## DB�� �������
            If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                Res = SaveTransData(gRow, spdOrder)

                If Res = -1 Then
                    '-- ���� ����
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "�������", gRow, colSTATE
                Else
                    '-- ���� ����
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX

                          SQL = "Update PATRESULT Set                                                               " & vbCrLf
                    SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                    SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                    SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                    SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                    SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                    If DBExec(AdoCn_Local, SQL) Then
                        '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_URINSCAN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_AVL9180()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    Dim Pos             As Integer
        
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
                    
                    '-- �˻���ó�� ���μ���
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
                
                .spdResult.RowHeight(-1) = 14
                
                '## DB�� �������
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)
    
                    If Res = -1 Then
                        '-- ���� ����
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "�������", gRow, colSTATE
                    Else
                        '-- ���� ����
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX
    
                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_AVL9180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_PATHFAST()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                    
                    '-- �������
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
                    
                    '-- ���ȯ������
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
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_PATHFAST" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_VISION()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                '-- 18�� ���
                strResult = mGetP(strRcvBuf, 10, vbTab)
                strIntResult = mGetP(strRcvBuf, 10, vbTab)
                'strResult = mGetP(strRcvBuf, 11, vbTab)

                '-- �������
                With mResult
                    .BarNo = strBarno
                    .Seq = strSeq
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- ���ȯ������
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                strState = "O"
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                        
                '-- �˻���ó�� ���μ���
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
                
                .spdResult.RowHeight(-1) = 15

                '## DB�� �������
                If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                    Res = SaveTransData(gRow, spdOrder)

                    If Res = -1 Then
                        '-- ���� ����
                        SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                        SetText .spdOrder, "�������", gRow, colSTATE
                    Else
                        '-- ���� ����
                        SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                        SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                        SetText .spdOrder, "0", gRow, colCHECKBOX

                              SQL = "Update PATRESULT Set                                                               " & vbCrLf
                        SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                        SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                        SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                        SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                        SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_TCPRcvData_KLITE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ISMART30()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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
                    strBarno = Trim(mGetP(strRcvBuf, 3, "|"))

                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_YUMIZEN()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    
    Dim intCnt          As Integer  '��� Frame ����
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
    
    '���� ����
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

                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_ISMART30" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_STAGO()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
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
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .Seq = strSeq
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
                        Case "F"    '## ����
                            strIntResult = strIntResult
                        Case "I"    '## ����
                            Select Case Mid$(strIntResult, 1, 1)
                                Case "N":   strResult = "Negative"
                                Case "G":   strResult = "GRAYZONE"
                                Case "R":   strResult = "Positive"
                                Case "P":   strResult = "Positive"
                            End Select
                    End Select
                        
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_STAGO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ACCESS2()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
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
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
'                        Case "F"    '## ����
'                            strResult = strIntResult
'                        Case "I"    '## ����
'                            Select Case Mid$(strIntResult, 1, 1)
'                                Case "N":   strResult = "Negative"
'                                Case "G":   strResult = "GRAYZONE"
'                                Case "R":   strResult = "Positive"
'                                Case "P":   strResult = "Positive"
'                            End Select
'                    End Select
                        
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                Case "L"
                
                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_ACCESS2" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_UROMETER720()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    strRData = Split(RcvBuffer, vbCrLf)
    
    With frmInterface
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "A")

            Select Case intCnt
                Case 3
                    strSeq = Mid(strRcvBuf, 10)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))
                    
                    '-- �������
                    mResult.Seq = strSeq
                    mResult.BarNo = strSeq
                    With mResult
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    
                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    strResult = ""
                    strIntResult = ""
                    strResult = Mid(strRcvBuf, 8, 4) '-- ����
                    strResult = Trim(strResult)
                    strIntResult = strResult
                    
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        strIntResult = Trim(Mid(strRcvBuf, 4))  '-- ����
                        strIntResult = Replace(strIntResult, "mg/dl", "")
                        strIntResult = Replace(strIntResult, "RBC/ul", "")
                        strIntResult = Replace(strIntResult, "WBC/ul", "")
                        
                        strIntResult = Replace(strIntResult, "<", "")
                        strIntResult = Replace(strIntResult, ">", "")
                        strIntResult = Replace(strIntResult, "=", "")
                        strResult = strIntResult
                    End If
                    
                    '-- URO
'                    If strResult = "norm" Then
'                        strResult = "Negative"
'                    End If
'
'                    '-- NIT
'                    If strResult = "pos" Then
'                        strResult = "1+"
'                    End If
'
'                    Select Case Trim(strResult)
'                        Case "-":       strResult = "Negative"
'                        Case "+":       strResult = "Pos(1+)"
'                        Case "++":      strResult = "Pos(2+)"
'                        Case "+++":     strResult = "Pos(3+)"
'                        Case "++++":    strResult = "Pos(4+)"
'                        Case "+/-":     strResult = "Trace(��)"
'                    End Select
                    
                    '-- �˻縶���� ���� ��������
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
                    
                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
        
        '## DB�� �������
        If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set                                                               " & vbCrLf
                SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_UROMETER720" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HORIBA()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
    Dim strCREA         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
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
                    
                    '-- �������
                    mOrder.Seq = strSeq
                    
                    With mResult
                        .BarNo = strSeq
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
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
                    
                    '-- �˻���ó�� ���μ���
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
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_HORIBA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

'-- �����
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
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strAMRResult    As String   '������ ���(����)
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqNo        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strAbbrName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
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
    
    '���� ����
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
                    
                    '-- �������
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
                                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                
                
                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                
                        '-- �˻���ó�� ���μ���
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
                
                    .spdResult.RowHeight(-1) = 15
                    

                    '## DB�� �������
                    If cn_Server_Flag = True And gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set                                                               " & vbCrLf
                            SQL = SQL & "       SENDFLAG = '2'                                                              " & vbCrLf
                            SQL = SQL & "     , SENDDATE = '" & Format(Now, "YYYYMMDD") & "'                                " & vbCrLf
                            SQL = SQL & " Where EQUIPNO  = '" & gHOSP.HOSPCD & "'                                           " & vbCrLf
                            SQL = SQL & "   And EXAMDATE = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And BARCODE  = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "'             " & vbCrLf
                            SQL = SQL & "   And SAVESEQ  = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SerialRcvData_AU480" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

Private Sub SendOrder_STAGO()
    Dim strOutput   As String     '�۽��� ������

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            
            '## �������� ������ �Ǵ��Ͽ� SndPhase����
            If mOrder.NoOrder = True Then
                '## ���������� ���°��
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
            '## ���� ������
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
            '## ���� ���ڿ��� ������
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_THUNDERBOLT()
    Dim strOutput   As String     '�۽��� ������
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
                If GetText(spdOrder, intRow, colCHECKBOX) = "1" And GetText(spdOrder, intRow, colSTATE) = "�����غ�" Then
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
                            Call SetText(spdOrder, "��������", intRow, colSTATE)
                            
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_XN1000()
    Dim strOutput   As String     '�۽��� ������

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
                '## ���� ������
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
                '## ���� ���ڿ��� ������
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800_ASTM()
    Dim strOutput   As String     '�۽��� ������

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
                '## ���� ������
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
                '## ���� ���ڿ��� ������
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_CA800()
    Dim strOutput   As String     '�۽��� ������

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
                '## ���� ������
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
                '## ���� ���ڿ��� ������
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_INDIKO()
    Dim strOutput   As String     '�۽��� ������

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||1^Analyzer 1^5.0|||||||P" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|||||||U||||||||||||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "||" & "|R|" & Format(Now, "yyyymmddhhmmss") & "||||||||||||||||||" & vbCr & ETX
                intSndPhase = 4
            Else
                '## ���� ������
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "||||||||||||||||||"
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ���� ���ڿ��� ������
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
            strOutput = intFrameNo & "L|1|F" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendData(EOT)
            intFrameNo = 1
            Call SetText(spdOrder, "��������", gRow, colSTATE)
            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_INDIKO_TCP()
    Dim strOutput   As String     '�۽��� ������

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||1^Analyzer 1^5.0|||||||P" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1|||||||U||||||||||||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "||" & "|R|" & Format(Now, "yyyymmddhhmmss") & "||||||||||||||||||" & vbCr & ETX
                intSndPhase = 4
            Else
                '## ���� ������
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R|" & Format(Now, "yyyymmddhhmmss") & "||||||||||||||||||"
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ���� ���ڿ��� ������
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
            strOutput = intFrameNo & "L|1|F" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            Call SendWSckData(EOT)
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendWSckData(strOutput)

End Sub

Private Sub SendOrder_ACCESS2()
    Dim strOutput   As String     '�۽��� ������
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
            '## ���� ������
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
            '## ���� ���ڿ��� ������
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
    Dim strOutput   As String     '�۽��� ������
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
                '## ���� ������
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
                '## ���� ���ڿ��� ������
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
                Call SetText(spdOrder, "��������", mOrder.DestRow, colSTATE)
                
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


Private Sub SendOrder_YUMIZEN()
    Dim strOutput   As String     '�۽��� ������
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
            '## ���� ������
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
            '## ���� ���ڿ��� ������
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
                        Call SetText(spdOrder, "��������", intRow, colSTATE)
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


Private Sub Phase_Serial_INDIKO()
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

                        If strState = "Q" Then
                            Call SendData(ENQ)
                            strState = ""

                            intSndPhase = 1
                            intFrameNo = 1
                            intPhase = 1
                        Else
                            Call SendData(ACK)
                            intPhase = 2
                        End If
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_INDIKO
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
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        intPhase = 1
                        Call SerialRcvData_INDIKO

                        'tmrQ.Interval = 200
                        'tmrQ.Enabled = True

                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                End Select
        End Select
    Next i
            
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'
'        Select Case intPhase
'            Case 1
'                Select Case Asc(BufChar)
'                    Case 5      'ENQ
'                        intPhase = 2
'                        RstEnd = "Y"
'                        bSTXChk = False
'                        bEndChk = True
'                        Call SendData(ACK)
'                    Case Else
'                        intPhase = 1
'                End Select
'            Case 2
'                Select Case Asc(BufChar)
'                    Case 2      'STX
'                        If bEndChk = True Then
'                            RcvBuffer = ""
'                        Else
'                            bSTXChk = True
'                        End If
'                        bEndChk = True
'                    Case 10     'LF
'                        If bEndChk = True Then
'                            Call SerialRcvData_INDIKO
'                            RcvBuffer = ""
'                        End If
'                        Call SendData(ACK)
'                    Case 13     'CR
'                        If bEndChk = True Then
'                            Call SerialRcvData_INDIKO
'                            RcvBuffer = ""
'                        End If
'                    Case 4      'EOT
'                        If strState = "Q" Then
'                            Call SendData(ENQ)
'                            intSndPhase = 1
'                        End If
'                        intPhase = 3
'                    Case 5      'ENQ
'                        bSTXChk = True
'                        bEndChk = True
'                        Call SendData(ACK)
'                    Case 21     'NAK
'                        Call SerialRcvData_INDIKO
'                        intSndPhase = 1
'                        intFrameNo = 1
'                        Call SendData(ENQ)
'                    Case 23     ' ETB
'                        bEndChk = False
'                    Case Else
'                        If bEndChk = True Then
'                            If bSTXChk = True Then
'                                bSTXChk = False
'                            Else
'                                RcvBuffer = RcvBuffer & BufChar
'                            End If
'                        End If
'                End Select
'
'            Case 3
'                Select Case Asc(BufChar)
'                    Case 6      'ACK
'                        If strState = "Q" Then
'                            Call SendOrder_INDIKO
'                        End If
'                    Case 5      'ENQ
'                        bSTXChk = False
'                        bEndChk = True
'                        Call SendData(ACK)
'                        intPhase = 2
'                    Case 21     'NAK
'                        intSndPhase = 1
'                        intFrameNo = 1
'                        Call SendData(ENQ)
'                        intPhase = 3
'                    Case 4      'EOT
'                        intPhase = 1
'
'                End Select
'        End Select
'    Next

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


Private Sub Phase_TCP_INDIKO()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1          '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intPhase = 2
                        Call SendWSckData(ACK)
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_INDIKO_TCP
                        End If
                End Select
            Case 2          '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendWSckData(ACK)
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
                        Call SendWSckData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        intPhase = 1
                        Call TCPRcvData_INDIKO
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendWSckData(ACK)
                        End If
                End Select
        End Select
    Next i

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

                'RcvBuffer = strBuffer

                Call TCPRcvData_VISION
                strBuffer = ""

            Case "CONNECT"
                strSendData = strSendAck & vbLf

                wSck.SendData strSendData
                SetRawData "[Tx]" & strSendData

            Case "RESULTS"
                '�����û
                strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))

                strNS = strRcvCnt
                strNE = mGetP(strBuffer, 4, vbTab)

                strNS = strNS - strNE
                strNE = strNS + strNE

                strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE & vbLf

                wSck.SendData strSendData
                SetRawData "[Tx]" & strSendData

                'Call WritePrivateProfileString("config", "LASTSEQ", strRcvCnt, App.PATH & "\Interface.ini")
                
                txtLastSeq.Text = strRcvCnt

                'blnResults = False
        End Select
    Next i


End Sub

Private Sub frmClear()
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
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
    lblSlipCd.Caption = gHOSP.PARTCD
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        If MsgBox("�������̽� ȭ���� �����ðڽ��ϱ�?", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
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
    
On Error GoTo ErrHandle
    
    'Me.Caption = gHOSP.MACHNM & Space$(5) & "�¢¢¢¢�     [���� �������̽�]     �¢¢¢¢�"
    'Me.Caption = gHOSP.HOSPNM & Space$(5) & gHOSP.MACHNM
    'Me.Caption = gHOSP.PARTNM & gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM

    '-- �ִ�ȭ
'''    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
'''        Me.WindowState = 2
'''    Else
'''        Me.WindowState = 0
'''        Me.TOP = gForm.TOP
'''        Me.LEFT = gForm.LEFT
'''        Me.WIDTH = gForm.WIDTH
'''        Me.HEIGHT = gForm.HEIGHT
'''    End If
    
    '-- �������� ǥ��
    If gHOSP.BARUSE = "Y" Then
        strIFStatus = "�� ���ڵ���"
    Else
        If gHOSP.RSTTYPE = "1" Then
            strIFStatus = "�� ���� ����"
        ElseIf gHOSP.RSTTYPE = "2" Then
            strIFStatus = "�� R/P ����"
        ElseIf gHOSP.RSTTYPE = "3" Then
            strIFStatus = "�� üũ��"
        End If
    End If
    strIFStatus = strIFStatus & IIf(gHOSP.SAVELIS = "Y", "  �� LIS���", "  �� �����")
    strIFStatus = strIFStatus & IIf(gHOSP.SAVEAUTO = "Y", "  �� �ڵ�����", "  �� ��������")
    lblIFStatus.Caption = strIFStatus
    
    '-- ��Ż��� ǥ��
    lblComStatus.Caption = ""
    
    '-- ��ũ��ȸ ��ġ
    If gWORKPOS = "M" Then
        spdWork.Visible = True
        fraWorkInfo.Visible = True
        'cmdView.Visible = True
    Else
        spdWork.Visible = False
        fraWorkInfo.Visible = False
        'cmdView.Visible = False
    End If

    '-- �ǻ���� ��� 'XML ����' ��ư ���̰���.
    If UCase(gEMR) = "UBCARE" Then
        cmdXML.Visible = True
    Else
        cmdXML.Visible = False
    End If
    
    '-- ��ź��� �ʱ�ȭ
    Call CtlInitializing

    '-- �� �ʱ�ȭ
    Call frmClear
    
    '-- �޴� ����
    Call SetMenu

    '-- �÷��������
    'Call SetColumnHeader(spdOrder)

    '-- �÷����̱⼳��
    Call SetColumnView(spdOrder)
    
    '-- �÷����̱⼳��
    Call SetColumnView(spdWork)
    
    '-- �÷����̱⼳��
    Call SetColumnViewResult(spdResult)
    
    '-- ��ũ����Ʈ �׸� ����
    For intCol = 1 To colSTATE
        spdWork.Col = intCol
        'If intCol = colHOSPDATE Or intCol = colBARCODE Or intCol = colPNAME Or intCol = colPID Or intCol = colCHECKBOX Then
        If intCol = colPNAME Or intCol = colCHECKBOX Then
            spdWork.ColHidden = False
        Else
            spdWork.ColHidden = True
        End If
    Next
    
'''    '-- �˻��� �׸� ����
'''    For intCol = 1 To colRPREVRESULT
'''        spdResult.Col = intCol
'''        If intCol = colRTESTNM Or intCol = colRLISRESULT Or intCol = colRPREVRESULT Then
'''            spdResult.ColHidden = False
'''        Else
'''            spdResult.ColHidden = True
'''        End If
'''    Next
    
    '-- �˻縶�������� gArrEQP(,) �� ���
    Call GetTestList

    '-- �˻��ڵ� gAllTestCd �� ���
    Call GetTestCodeList

    '-- �˻縶���͸� gArrEQPNm �� ���
    'Call GetTestListName

    '-- �˻�� ���̱�
    Call SetExamCode(spdOrder)

    '-- ��ſ���
    Call OpenCommunication
    
    '���� �ʱ�ȭ(E-170/H-7600)
    RstEnd = "Y"
    bSTXChk = False
    bEndChk = True
    
    pDel = False

    imgNet1.ZOrder 0
    tmrDBConn.Interval = 1000
    tmrDBConn.Enabled = True
    
    tmrQ = False
    
    '-- ������� ����
    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gHOSP.SAVEDAY + "���� ����Ÿ�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    

    If gHOSP.DBCONCHK = "Y" Then
        tmrConn.Interval = 60000
        tmrConn.Enabled = True
    Else
        tmrConn.Enabled = False
    End If
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("��Ʈ ��ȣ�� �߸��Ǿ����ϴ�." & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "��Ʈ �������"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            
            Resume Next
        Else
            End
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
            lblComStatus.Caption = "COM" & comEqp.CommPort & "��Ʈ ���Ἲ��"
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgOn.ZOrder 0

        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & "��Ʈ �������"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgOff.ZOrder 0
        
           ' imgCom.Picture = imlStatus.ListImages("OFF").ExtractIcon
        
        End If
        
    ElseIf gComm.COMTYPE = "2" Then
        'lblComStatus.Left = imgPort.Left + 500
        'lblComStatus.Width = 6000
        If gComm.TCPTYPE = "SERVER" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen
            
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ������.."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            'imgSend.Visible = False
            'imgReceive.Visible = False
            'lblSend.Visible = False
            'lblRcv.Visible = False
            imgOff.ZOrder 0

        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
            
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ������..."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
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

    'Exit Sub
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'Me.TOP = 0
    MDIIF.cmdNode.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT

    If gWORKPOS = "M" Then
        spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - fraWorkInfo.HEIGHT - picBottom.HEIGHT - 160
        
        If spdResult.Visible = True Then
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = picTop.TOP + picTop.HEIGHT + 40
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
            
            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
        Else
            spdOrder.LEFT = spdWork.WIDTH + 100
            spdOrder.TOP = picTop.TOP + picTop.HEIGHT + 40
            spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - 200
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
        End If
    Else
        If spdResult.Visible = True Then
            spdOrder.LEFT = picTop.LEFT + 40
            spdOrder.TOP = picTop.TOP + picTop.HEIGHT + 40
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200
            
            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
        Else
            spdOrder.LEFT = picTop.LEFT + 40
            spdOrder.TOP = picTop.TOP + picTop.HEIGHT + 40
            spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picTop.HEIGHT - picBottom.HEIGHT - 100
            spdOrder.WIDTH = Me.ScaleWidth - 200
            
            spdResult.TOP = spdOrder.TOP
            spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
            spdResult.HEIGHT = spdOrder.HEIGHT
        End If
    End If
    
'''    If Me.WindowState = 2 Then
'''        'gForm.MAXYN = True
'''        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
'''    Else
'''        If Me.TOP < 0 Then
'''            Me.TOP = 0
'''        End If
'''        'gForm.MAXYN = False
'''        gForm.TOP = Me.TOP
'''        gForm.LEFT = Me.LEFT
'''        gForm.WIDTH = Me.WIDTH
'''        gForm.HEIGHT = Me.HEIGHT
'''
'''        Call WritePrivateProfileString("FORM", "MAXYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
'''        Call WritePrivateProfileString("FORM", "TOP", gForm.TOP, App.PATH & "\INI\" & gMACH & ".ini")
'''        Call WritePrivateProfileString("FORM", "LEFT", gForm.LEFT, App.PATH & "\INI\" & gMACH & ".ini")
'''        Call WritePrivateProfileString("FORM", "WIDTH", gForm.WIDTH, App.PATH & "\INI\" & gMACH & ".ini")
'''        Call WritePrivateProfileString("FORM", "HEIGHT", gForm.HEIGHT, App.PATH & "\INI\" & gMACH & ".ini")
'''
'''    End If
    
End Sub

'�������̽� ȯ�ڼ��ý� ������ �˻��׸�/��������ֱ�
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
    SQL = SQL & "SELECT DISTINCT SEQNO, EQUIPCODE, EXAMNAME, EXAMCODE, EQUIPRESULT, RESULT, PREVRESULT, REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE = '" & strBarno & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmInterface.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                If AdoRs_Local.Fields("EXAMCODE").Value & "" = "" Then
                    Call SetText(frmInterface.spdResult, "0", intRow, colCHECKBOX)
                Else
                    Call SetText(frmInterface.spdResult, "1", intRow, colCHECKBOX)
                End If
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EQUIPCODE").Value & "", intRow, colRCHANNEL)
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("EQUIPRESULT").Value & "", intRow, colRMACHRESULT)
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                If AdoRs_Local.Fields("REFJUDGE").Value & "" = "H" Then
                    .ForeColor = vbRed
                    .FontBold = True
                ElseIf AdoRs_Local.Fields("REFJUDGE").Value & "" = "L" Then
                    .ForeColor = vbBlue
                    .FontBold = True
                Else
                    .ForeColor = vbBlack
                    .FontBold = False
                End If
                Call SetText(frmInterface.spdResult, AdoRs_Local.Fields("PREVRESULT").Value & "", intRow, colRPREVRESULT)
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
        GetPatTRestResult = 1
    End If

    AdoRs_Local.Close

Exit Function

ErrHandle:
    GetPatTRestResult = -1

    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
        lblComStatus.Caption = "COM" & comEqp.CommPort & "��Ʈ ���Ἲ��"
        
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "��Ʈ �������"
        
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    End If

End Sub



Private Sub lblSlipCd_DblClick()
    Dim strSlipCd   As String
    
    strSlipCd = InputBox("SLIP �ڵ��Է�", "SLIP CD", lblSlipCd.Caption)
        
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
    
    '-- ����
'    If Row = 0 Then
'        '-- ���� �߰�
'        Exit Sub
'    End If
    
    
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
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, Row, colCHECKBOX) = "1" Then
            Call SetText(spdOrder, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdOrder, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
    
    If Row = 0 Then
        Exit Sub
    End If
    
    '-- ȯ������ǥ��
    lblBarcode.Caption = GetText(spdOrder, Row, colBARCODE)
    
    strPName = GetText(spdOrder, Row, colPNAME)
    strPSex = GetText(spdOrder, Row, colPSEX)
    strPSex = IIf(strPSex = "", "-", strPSex)
    strPAge = GetText(spdOrder, Row, colPAGE)
    strPAge = IIf(strPAge = "", "-", strPAge)
    
    lblPatNm.Caption = strPName & Space(1) & strPSex & "/" & strPAge
    
    lblStatus.Caption = IIf(GetText(spdOrder, Row, colSTATE) = "", "�˻��غ�", GetText(spdOrder, Row, colSTATE))
    
    If chkAdd.Value = "1" Then
        txtOldBarNum.Text = GetText(spdOrder, Row, colBARCODE)
    Else
        txtOldBarNum.Text = ""
    End If
    
    '-- ���ǥ��
    If GetPatTRestResult(Row) = -1 Then
        '������� ������� �˻�� �����ֱ�
        spdResult.MaxRows = 0
        strItems = ""
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '��
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    strItems = strItems & GetText(spdOrder, 0, intCol) & "/"
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 15
                End If
            Next
        End With
    End If

    lblRow.Caption = Row
    
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
                MsgBox "�Է��� ���ڵ忡�� ȯ�������� ã�� ���߽��ϴ�." & vbNewLine & " ���ڵ� ��ȣ�� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
            Else
                '��������
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
                    '-- ����
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "���ڸ� �Է��� �����մϴ�"
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
        
        If MsgBox(strNewBarNo & " �� ����ðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
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






Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
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
            '-- ���� �߰�
            Call SetSpreadSort(spdWork, 0)
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
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
            
            For i = colEXAMDATE To colSTATE ' colCHECKBOX To colSTATE
                Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
            Next
            
            varItems = GetText(spdWork, intWRow, colITEMS)
            varItems = Split(varItems, "/")
            For intItems = 0 To UBound(varItems)
                For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                    .Row = 0
                    .Col = intOCol
                    If varItems(intItems) = Trim(.Text) Then
                        .Row = intRow
                        Call SetText(spdOrder, "��", intRow, intOCol)
                    End If
                Next
            Next
            
            
            
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
            .RowHeight(-1) = 15
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
        If MsgBox(strPatName & " �� ����ðڽ��ϱ�?", vbCritical + vbYesNo, "�˸�") = vbNo Then
            Exit Sub
        End If
        '��������
        SQL = ""
        SQL = SQL & "DELETE FROM UB_PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdWork, sRow, colBARCODE)) & "'" & vbCr
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- ����
            DeleteRow spdWork, sRow, sRow
            spdWork.MaxRows = spdWork.MaxRows - 1
        End If
    End If
End Sub

Private Sub tmrConn_Timer()
    Dim sqlRet          As Long
    Dim RS          As ADODB.Recordset
    
On Error GoTo ErrHandle
    If DbConnect_SQL = True Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("Select sysdate from DUAL", sqlRet)
        RS.Close
        
        ''Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmInterface.lstComStatus)
    End If
Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    lblDBStatus.Caption = "�����ͺ��̽� �������"
'    frmErrMsg.Show
    
End Sub

Private Sub tmrDBConn_Timer()

    DoEvents

    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
'    MDIIF.lblTestDate.ForeColor = RGB(255, Int((255 * Rnd) + 1), 0)
    
    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If
    
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
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
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
                    MsgBox "�Է��� ���ڵ忡�� ȯ�������� ã�� ���߽��ϴ�." & vbNewLine & " ���ڵ� ��ȣ�� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
                Else
                    '��������
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
                        '-- ����
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
                    MsgBox "�Է��� ���ڵ忡�� ȯ�������� ã�� ���߽��ϴ�." & vbNewLine & " ���ڵ� ��ȣ�� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
                Else
                    '��������
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
                        '-- ����
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
            MsgBox "���ڸ� �Է��� �����մϴ�"
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
            MsgBox "���ڸ� �Է��� �����մϴ�"
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

Private Sub txtSeqNo_KeyPress(KeyAscii As Integer)
    Dim intSeq  As Integer
    Dim intRow  As Integer
                
    
    If KeyAscii = vbKeyReturn Then
        intSeq = txtSeqNo.Text
        
        If Not IsNumeric(intSeq) Then
            MsgBox "���ڸ� �Է��� �����մϴ�"
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

Private Sub wSCK_Close()
        
    If gComm.TCPTYPE = "SERVER" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
        imgOn.ZOrder 0
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
        imgOn.ZOrder 0
    End If

End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "SERVER" Then
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
            imgOn.ZOrder 0
        Else
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
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

