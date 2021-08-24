VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBarConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Barcode라벨 출력양식 설정"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   10860
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Barcode Setting"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   390
      TabIndex        =   16
      Top             =   2655
      Width           =   10050
      Begin VB.ComboBox cboStyle 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":0000
         Left            =   7725
         List            =   "frmBarConfigLEO60D.frx":000D
         Style           =   2  '드롭다운 목록
         TabIndex        =   24
         Top             =   240
         Width           =   2100
      End
      Begin VB.Frame fraDetail 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         Height          =   435
         Index           =   14
         Left            =   1695
         TabIndex        =   18
         Top             =   165
         Width           =   8145
         Begin VB.TextBox txtBarHeight 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   0  '없음
            Height          =   210
            Left            =   4800
            TabIndex        =   23
            Text            =   "1234"
            Top             =   120
            Width           =   480
         End
         Begin VB.TextBox txtBarPosX 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   0  '없음
            Height          =   210
            Left            =   870
            TabIndex        =   21
            Text            =   "1234"
            Top             =   135
            Width           =   480
         End
         Begin VB.TextBox txtBarPosY 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   0  '없음
            Height          =   210
            Left            =   1695
            TabIndex        =   20
            Text            =   "1234"
            Top             =   135
            Width           =   480
         End
         Begin VB.TextBox txtBarLength 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   0  '없음
            Height          =   210
            Left            =   2970
            TabIndex        =   19
            Text            =   "12"
            Top             =   135
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   2
            Height          =   255
            Index           =   4
            Left            =   4770
            Shape           =   4  '둥근 사각형
            Top             =   105
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "위치 : X            Y             길이 :           Barcode높이 :            Style : "
            Height          =   180
            Index           =   14
            Left            =   120
            TabIndex        =   22
            Top             =   150
            Width           =   5925
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   2
            Height          =   255
            Index           =   48
            Left            =   840
            Shape           =   4  '둥근 사각형
            Top             =   120
            Width           =   555
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   2
            Height          =   255
            Index           =   47
            Left            =   1665
            Shape           =   4  '둥근 사각형
            Top             =   120
            Width           =   555
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   2
            Height          =   255
            Index           =   46
            Left            =   2940
            Shape           =   4  '둥근 사각형
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.CheckBox chkBarcode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출력여부"
         Height          =   240
         Left            =   435
         TabIndex        =   17
         Tag             =   "BAR_BARCD"
         Top             =   285
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "About Label"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   390
      TabIndex        =   5
      Top             =   915
      Width           =   10050
      Begin VB.ComboBox cboClientPort 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":002F
         Left            =   4305
         List            =   "frmBarConfigLEO60D.frx":003F
         Style           =   2  '드롭다운 목록
         TabIndex        =   48
         Tag             =   "BAR_PORT"
         Top             =   405
         Width           =   1170
      End
      Begin VB.TextBox txtGapLen 
         BorderStyle     =   0  '없음
         Height          =   225
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "BAR_GAP"
         Top             =   1215
         Width           =   1095
      End
      Begin VB.ComboBox cboPort 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":005F
         Left            =   1590
         List            =   "frmBarConfigLEO60D.frx":006F
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Tag             =   "BAR_PORT"
         Top             =   405
         Width           =   1170
      End
      Begin VB.TextBox txtTotLen 
         BorderStyle     =   0  '없음
         Height          =   225
         Left            =   1620
         TabIndex        =   8
         Tag             =   "BAR_TOTLEN"
         Top             =   1230
         Width           =   1080
      End
      Begin VB.TextBox txtLength 
         BorderStyle     =   0  '없음
         Height          =   225
         Left            =   4350
         TabIndex        =   7
         Tag             =   "BAR_LEN"
         Top             =   825
         Width           =   1110
      End
      Begin VB.TextBox txtWidth 
         BorderStyle     =   0  '없음
         Height          =   225
         Left            =   1620
         TabIndex        =   6
         Tag             =   "BAR_WIDTH"
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Client Port  : "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   8
         Left            =   2925
         TabIndex        =   49
         Top             =   450
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   1545
         Left            =   5745
         Picture         =   "frmBarConfigLEO60D.frx":008F
         Top             =   120
         Width           =   4170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   5700
         X2              =   5700
         Y1              =   180
         Y2              =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   5685
         X2              =   5685
         Y1              =   180
         Y2              =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   270
         Index           =   0
         Left            =   4320
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Gap Length   : "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   5
         Left            =   2940
         TabIndex        =   15
         Top             =   1230
         Width           =   1350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   270
         Index           =   3
         Left            =   1590
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   270
         Index           =   2
         Left            =   4320
         Top             =   810
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FBC2A4&
         BorderWidth     =   2
         Height          =   270
         Index           =   1
         Left            =   1590
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Total Length :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   3
         Left            =   255
         TabIndex        =   12
         Top             =   1245
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Label Length : "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   2
         Left            =   2940
         TabIndex        =   11
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Label Width  :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   1
         Left            =   255
         TabIndex        =   10
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Serial Port  : "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00172C2D&
         Height          =   180
         Index           =   4
         Left            =   255
         TabIndex        =   9
         Top             =   450
         Width           =   1350
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   900
      Left            =   390
      TabIndex        =   0
      Top             =   0
      Width           =   10050
      Begin VB.ComboBox cboClientKind 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":0DC4
         Left            =   5745
         List            =   "frmBarConfigLEO60D.frx":0DDD
         Style           =   2  '드롭다운 목록
         TabIndex        =   47
         Top             =   525
         Width           =   1800
      End
      Begin VB.ComboBox cboBarKind 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":0E25
         Left            =   5745
         List            =   "frmBarConfigLEO60D.frx":0E3E
         Style           =   2  '드롭다운 목록
         TabIndex        =   44
         Top             =   195
         Width           =   1800
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종료"
         Height          =   510
         Left            =   8820
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "저장"
         Height          =   510
         Left            =   7725
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   240
         Width           =   1080
      End
      Begin VB.ComboBox cboProject 
         Height          =   300
         ItemData        =   "frmBarConfigLEO60D.frx":0E87
         Left            =   1395
         List            =   "frmBarConfigLEO60D.frx":0E94
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   195
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Client 기종  :"
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
         Index           =   7
         Left            =   4320
         TabIndex        =   46
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Barcode기종:"
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
         Index           =   6
         Left            =   4320
         TabIndex        =   45
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Label종류  : "
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
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detail Information"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5850
      Left            =   390
      TabIndex        =   25
      Top             =   3375
      Width           =   10050
      Begin VB.Frame fraCollectLabel 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5550
         Left            =   210
         TabIndex        =   26
         Top             =   225
         Width           =   9765
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   1020
            Left            =   0
            ScaleHeight     =   960
            ScaleWidth      =   9600
            TabIndex        =   28
            Top             =   0
            Width           =   9660
            Begin VB.CheckBox chkErFg 
               BackColor       =   &H00E0E0E0&
               Caption         =   "응급실환자 표시 (Reverse)"
               Height          =   240
               Left            =   210
               TabIndex        =   39
               Top             =   675
               Width           =   2580
            End
            Begin VB.CheckBox chkStat 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Stat 처방 표시 (Line)"
               Height          =   240
               Left            =   210
               TabIndex        =   38
               Top             =   360
               Width           =   2115
            End
            Begin VB.CheckBox chkAccCheck 
               BackColor       =   &H00E0E0E0&
               Caption         =   "접수여부 Check"
               Height          =   240
               Left            =   210
               TabIndex        =   37
               Top             =   60
               Width           =   1890
            End
            Begin VB.TextBox txtERcode 
               Alignment       =   2  '가운데 맞춤
               BorderStyle     =   0  '없음
               Height          =   210
               Left            =   3990
               TabIndex        =   36
               Text            =   "1234"
               Top             =   675
               Width           =   840
            End
            Begin VB.ComboBox cboReverse 
               Height          =   300
               ItemData        =   "frmBarConfigLEO60D.frx":0EE5
               Left            =   6315
               List            =   "frmBarConfigLEO60D.frx":0F13
               Style           =   2  '드롭다운 목록
               TabIndex        =   35
               Top             =   630
               Width           =   1875
            End
            Begin VB.Frame fraLine 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  '없음
               Height          =   345
               Left            =   2910
               TabIndex        =   29
               Top             =   270
               Width           =   5430
               Begin VB.TextBox txtLineLength 
                  Alignment       =   2  '가운데 맞춤
                  BorderStyle     =   0  '없음
                  Height          =   210
                  Left            =   3090
                  TabIndex        =   33
                  Text            =   "12"
                  Top             =   90
                  Width           =   375
               End
               Begin VB.TextBox txtStatPosY 
                  Alignment       =   2  '가운데 맞춤
                  BorderStyle     =   0  '없음
                  Height          =   210
                  Left            =   1620
                  MaxLength       =   4
                  TabIndex        =   32
                  Text            =   "1234"
                  Top             =   90
                  Width           =   480
               End
               Begin VB.TextBox txtStatPosX 
                  Alignment       =   2  '가운데 맞춤
                  BorderStyle     =   0  '없음
                  Height          =   210
                  Left            =   795
                  MaxLength       =   4
                  TabIndex        =   31
                  Text            =   "1234"
                  Top             =   90
                  Width           =   480
               End
               Begin VB.TextBox txtLineWidth 
                  Alignment       =   2  '가운데 맞춤
                  BorderStyle     =   0  '없음
                  Height          =   210
                  Left            =   4695
                  TabIndex        =   30
                  Text            =   "1234"
                  Top             =   90
                  Width           =   480
               End
               Begin VB.Shape Shape2 
                  BorderColor     =   &H00C0C0C0&
                  Height          =   315
                  Left            =   15
                  Shape           =   4  '둥근 사각형
                  Top             =   30
                  Width           =   5250
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00A8A8A8&
                  BorderWidth     =   2
                  Height          =   255
                  Index           =   8
                  Left            =   3060
                  Shape           =   4  '둥근 사각형
                  Top             =   75
                  Width           =   435
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00A8A8A8&
                  BorderWidth     =   2
                  Height          =   255
                  Index           =   7
                  Left            =   1590
                  Shape           =   4  '둥근 사각형
                  Top             =   75
                  Width           =   555
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00A8A8A8&
                  BorderWidth     =   2
                  Height          =   255
                  Index           =   6
                  Left            =   765
                  Shape           =   4  '둥근 사각형
                  Top             =   75
                  Width           =   555
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '투명
                  Caption         =   "위치 : X            Y                길이 :            Line두께 : "
                  Height          =   240
                  Index           =   0
                  Left            =   45
                  TabIndex        =   34
                  Top             =   105
                  Width           =   4560
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00A8A8A8&
                  BorderWidth     =   2
                  Height          =   255
                  Index           =   5
                  Left            =   4665
                  Shape           =   4  '둥근 사각형
                  Top             =   75
                  Width           =   555
               End
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "Reverse항목"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   5145
               TabIndex        =   41
               Top             =   690
               Width           =   1050
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "응급실코드"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   2955
               TabIndex        =   40
               Top             =   690
               Width           =   900
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00A8A8A8&
               BorderWidth     =   2
               Height          =   255
               Index           =   9
               Left            =   3960
               Shape           =   4  '둥근 사각형
               Top             =   660
               Width           =   915
            End
         End
         Begin FPSpread.vaSpread tblDetails 
            Height          =   4500
            Left            =   0
            TabIndex        =   27
            Top             =   1035
            Width           =   9645
            _Version        =   196608
            _ExtentX        =   17013
            _ExtentY        =   7938
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   14
            ScrollBars      =   0
            ShadowColor     =   15530489
            ShadowDark      =   13753559
            SpreadDesigner  =   "frmBarConfigLEO60D.frx":0FC3
            TextTip         =   4
         End
      End
      Begin VB.Frame fraBloodLabel 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   210
         TabIndex        =   42
         Top             =   225
         Width           =   9765
         Begin FPSpread.vaSpread tblBldDetails 
            Height          =   3150
            Left            =   30
            TabIndex        =   43
            Top             =   285
            Width           =   9645
            _Version        =   196608
            _ExtentX        =   17013
            _ExtentY        =   5556
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   8
            ScrollBars      =   0
            ShadowColor     =   15530489
            ShadowDark      =   13753559
            SpreadDesigner  =   "frmBarConfigLEO60D.frx":190A
            TextTip         =   4
         End
      End
   End
End
Attribute VB_Name = "frmBarConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objBarConfig    As New clsBarConfig
Private objLabelConfig  As New clsBarConfig
Private mudtBarcode     As tpBarcode
Private mudtStatInfo    As tpStatInfo
Private mudtBarData(14) As tpBarData
Private strFontList     As String
Private mvarProjectId   As String

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Private Sub cboBarKind_Click()
    Dim strPrjNm    As String
    Dim strBarKind  As String
    Dim strGetKey   As String
    Dim i           As Long
    
    If cboProject.ListIndex < 0 Then Exit Sub
    If cboBarKind.ListIndex < 1 Then Exit Sub
    
    strPrjNm = Mid(cboProject.Text, 1, 3)
    strBarKind = cboBarKind.ListIndex
    
    mvarProjectId = strPrjNm
    
    'Client Barcode information

    strGetKey = strPrjNm & COL_DIV & strBarKind
    
    If cboProject.ItemData(cboProject.ListIndex) < 4 Then
        fraCollectLabel.Visible = True
        fraBloodLabel.Visible = False
        DoEvents
        Call GetCollectionLabelInfo(strGetKey) '(strPrjNm)
    Else
        fraCollectLabel.Visible = False
        fraBloodLabel.Visible = True
        DoEvents
        Call GetBloodLabelInfo(strGetKey) '(strPrjNm)
    End If
    
End Sub

Private Sub cboProject_Click()
    Dim strPrjNm    As String
    Dim strBarKind  As String
    Dim strGetKey   As String
    Dim i           As Long
    
    If cboProject.ListIndex < 0 Then Exit Sub
    
    strPrjNm = Mid(cboProject.Text, 1, 3)
    mvarProjectId = strPrjNm
    
    'Client Barcode information

    Call GetClientBarInfo(strPrjNm, strBarKind)
    strGetKey = strPrjNm & COL_DIV & strBarKind
    
    If cboProject.ItemData(cboProject.ListIndex) < 4 Then
        fraCollectLabel.Visible = True
        fraBloodLabel.Visible = False
        DoEvents
        Call GetCollectionLabelInfo(strGetKey) '(strPrjNm)
    Else
        fraCollectLabel.Visible = False
        fraBloodLabel.Visible = True
        DoEvents
        Call GetBloodLabelInfo(strGetKey) '(strPrjNm)
    End If
    
End Sub

Private Sub GetCollectionLabelInfo(ByVal strPrjNm As String)
    
    Dim i   As Long
    
    Set objBarConfig = New clsBarConfig
    With objBarConfig
        Call objBarConfig.ReadConfig(strPrjNm)
        cboPort.ListIndex = .PortNo - 1
        
'        If .BarKind <> "" Then
'            cboBarKind.ListIndex = .BarKind
'        Else
'            cboBarKind.ListIndex = 0
'        End If
        If .MainBarKind <> "" Then
            cboBarKind.ListIndex = .MainBarKind
        Else
            cboBarKind.ListIndex = 0
        End If
        txtWidth.Text = .Width
        txtLength.Text = .Length
        txtTotLen.Text = .TotLength
        chkBarcode.Value = Val(.Barcode.PrtFg)
        txtBarPosX.Text = .Barcode.PosX
        txtBarPosY.Text = .Barcode.PosY
        txtBarLength.Text = .Barcode.Length
        txtBarHeight.Text = .Barcode.Height
        cboStyle.ListIndex = .Barcode.Style
        chkAccCheck.Value = Val(.AccCheck)
        chkStat.Value = Val(.StatInfo.PrtLineFg)
        chkErFg.Value = Val(.StatInfo.PrtReverseFg)
        txtStatPosX.Text = .StatInfo.PosX
        txtStatPosY.Text = .StatInfo.PosY
        txtLineLength.Text = .StatInfo.Length
        txtLineWidth.Text = .StatInfo.Width
        txtERcode.Text = .StatInfo.ErDeptCd
        cboReverse.ListIndex = .StatInfo.ReverseFld
    End With
    For i = 1 To tblDetails.MaxRows
        With objBarConfig.BarData(i)
            tblDetails.Row = i
            tblDetails.Col = 2: tblDetails.Value = .PrtFg
            tblDetails.Col = 3: tblDetails.Value = .PosX
            tblDetails.Col = 4: tblDetails.Value = .PosY
            tblDetails.Col = 5: tblDetails.Value = .Length
            tblDetails.Col = 6: tblDetails.Text = medGetP(strFontList, Val(.FontX) + 1, vbTab)
                                tblDetails.CellType = CellTypeStaticText
            tblDetails.Col = 7: tblDetails.Text = medGetP(strFontList, Val(.FontY) + 1, vbTab)
                                tblDetails.CellType = CellTypeStaticText
            tblDetails.Col = 8: tblDetails.Value = .BoldFg
            tblDetails.Col = 9: tblDetails.Value = .ReverseFg
        End With
    Next
    tblDetails.Refresh
    
End Sub

Private Sub GetBloodLabelInfo(ByVal strPrjNm As String)

    Dim i As Long
    
    Set objLabelConfig = New clsBarConfig
    With objLabelConfig
        Call objLabelConfig.ReadConfig(strPrjNm)
        
        cboPort.ListIndex = .PortNo - 1
        If .BarKind <> "" Then
            cboBarKind.ListIndex = .BarKind
        Else
            cboBarKind.ListIndex = 0
        End If
        txtWidth.Text = .Width
        txtLength.Text = .Length
        txtTotLen.Text = .TotLength
        chkBarcode.Value = Val(.Barcode.PrtFg)
        txtBarPosX.Text = .Barcode.PosX
        txtBarPosY.Text = .Barcode.PosY
        txtBarLength.Text = .Barcode.Length
        txtBarHeight.Text = .Barcode.Height
        cboStyle.ListIndex = .Barcode.Style
    End With
    For i = 1 To tblBldDetails.MaxRows
        With objLabelConfig.BarData(i)
            tblBldDetails.Row = i
           
            tblBldDetails.Col = 2: tblBldDetails.Value = .PrtFg
            tblBldDetails.Col = 3: tblBldDetails.Value = .PosX
            tblBldDetails.Col = 4: tblBldDetails.Value = .PosY
            tblBldDetails.Col = 5: tblBldDetails.Value = .Length
            tblBldDetails.Col = 6: tblBldDetails.Text = medGetP(strFontList, Val(.FontX) + 1, vbTab)
                                   tblBldDetails.CellType = CellTypeStaticText
            tblBldDetails.Col = 7: tblBldDetails.Text = medGetP(strFontList, Val(.FontY) + 1, vbTab)
                                   tblBldDetails.CellType = CellTypeStaticText
            tblBldDetails.Col = 8: tblBldDetails.Value = .BoldFg
            tblBldDetails.Col = 9: tblBldDetails.Value = .ReverseFg
        End With
    Next
    tblBldDetails.Refresh
    
End Sub
'--------------------------------------------------
' Client 별 바코드 정보
'   - 인수
'       - pPrjNm : ProjectID
'   - 결과
'       Client Barcode Port
'       Client Barcode Kind
'--------------------------------------------------
Private Sub GetClientBarInfo(ByVal pPrjNm As String, ByRef pBarKind As String)
    Dim objSQL  As clsBarSqlStmt
    Dim Rs      As Recordset
    Dim strPort As String
    Dim strKind As String
'    Dim strPath As String
    
'    strPath = INIPath
    
    If Dir(INIPath) = "" Then
JUMP:
        Set Rs = New Recordset
        Set objSQL = New clsBarSqlStmt
        Rs.Open objSQL.SQL_ReadConfig(pPrjNm), DBConn
        
        Do Until Rs.EOF
            Select Case Rs.Fields("cdval2").Value & ""
                Case BAR_KIND
                    strPort = Rs.Fields("field1").Value & ""
                Case BAR_PORT
                    strKind = Rs.Fields("field1").Value & ""
            End Select
            Rs.MoveNext
        Loop
        Set Rs = Nothing
        Set objSQL = Nothing
        
        Call medSetINI(pPrjNm, "PORT", strPort, INIPath)
        Call medSetINI(pPrjNm, "KIND", strKind, INIPath)
    Else
        If medGetINI(pPrjNm, "PORT", INIPath) = "" Then GoTo JUMP
    End If
    
    cboClientPort.ListIndex = Val(medGetINI(pPrjNm, "PORT", INIPath)) - 1
    cboClientKind.ListIndex = Val(medGetINI(pPrjNm, "KIND", INIPath))
    
    pBarKind = cboClientKind.ListIndex
    
End Sub


Private Sub cmdExit_Click()
    Unload Me
    Set frmBarConfig = Nothing
End Sub


Private Sub cmdSave_Click()
    Dim strPrjNm        As String   ' ProjectID
    Dim strSaveKey      As String   ' 저장Key
    Dim strMainKind     As String   ' 바코드 종료(Main)
    Dim strClientKind   As String   ' 바코드 종류(Client)
    Dim strMainPORT     As String   ' 바코드 포트(Main)
    Dim strClientPort   As String   ' 바코드 포트(Client)
    Dim i               As Long     '
    
    If cboBarKind.ListIndex = 0 Then
        MsgBox "바코드 기종을 선택하세요.", vbInformation + vbOKOnly, "바코드기종선택"
        Exit Sub
    End If
    
    strMainKind = cboBarKind.ListIndex
    strClientKind = cboClientKind.ListIndex
    strMainPORT = cboPort.ListIndex + 1
    strClientPort = cboClientPort.ListIndex + 1
    
    strPrjNm = Mid(cboProject.Text, 1, 3)
    strSaveKey = strPrjNm & COL_DIV & strMainKind
    
    With objBarConfig
        .PortNo = strMainPORT
        .Width = Format(txtWidth.Text, "0###")
        .Length = Format(txtLength.Text, "0###")
        .TotLength = Format(txtTotLen.Text, "0###")
        .BarKind = strMainKind
        mudtBarcode = .Barcode
        mudtBarcode.PrtFg = chkBarcode.Value
        mudtBarcode.PosX = Format(txtBarPosX.Text, "0###")
        mudtBarcode.PosY = Format(txtBarPosY.Text, "0###")
        mudtBarcode.Length = Format(txtBarLength.Text, "0#")
        mudtBarcode.Height = Format(txtBarHeight.Text, "0###")
        mudtBarcode.Style = Format(cboStyle.ListIndex, "0#")
        .Barcode = mudtBarcode
        
        If cboProject.ItemData(cboProject.ListIndex) < 4 Then
            .AccCheck = chkAccCheck.Value
            mudtStatInfo = .StatInfo
            mudtStatInfo.PrtLineFg = chkStat.Value
            mudtStatInfo.PrtReverseFg = chkErFg.Value
            mudtStatInfo.PosX = Format(txtStatPosX.Text, "0###")
            mudtStatInfo.PosY = Format(txtStatPosY.Text, "0###")
            mudtStatInfo.Length = Format(txtLineLength.Text, "0###")
            mudtStatInfo.Width = Format(txtLineWidth.Text, "0###")
            mudtStatInfo.ErDeptCd = txtERcode.Text
            mudtStatInfo.ReverseFld = cboReverse.ListIndex
            .StatInfo = mudtStatInfo
        End If
    End With
    
    If cboProject.ItemData(cboProject.ListIndex) < 4 Then
        Call SetDetailInfo(tblDetails)
    Else
        Call SetDetailInfo(tblBldDetails)
    End If

    ' 바코드의 기종과 PC별로 Port 를 INI로 가지고 있는다..
    Call medSetINI(strPrjNm, "PORT", strClientPort, INIPath)
    Call medSetINI(strPrjNm, "KIND", strClientKind, INIPath)
    
    If objBarConfig.SaveConfig(strSaveKey) Then     '    strPrjNm
        MsgBox "정상적으로 저장되었습니다.", vbInformation + vbOKOnly, "바코드 출력양식 설정"
        Call ClearRtn
    Else
        MsgBox "저장시 오류가 발생했습니다. 전산실로 문의바랍니다.", vbExclamation + vbOKOnly, "오류발생"
    End If
End Sub

Private Sub SetDetailInfo(ByVal objTable As Object)

    Dim i As Long
    
    For i = 1 To objTable.MaxRows
        mudtBarData(i) = objBarConfig.BarData(i)
        With mudtBarData(i)
            objTable.Row = i
            objTable.Col = 2: .PrtFg = objTable.Value
            objTable.Col = 3: .PosX = Format(objTable.Value, "0###")
            objTable.Col = 4: .PosY = Format(objTable.Value, "0###")
            objTable.Col = 5: .Length = Format(objTable.Value, "0#")
            objTable.Col = 6: .FontX = medGetP(objTable.Text, 1, ".")
            objTable.Col = 7: .FontY = medGetP(objTable.Text, 1, ".")
            
            objTable.Col = 8: .BoldFg = objTable.Value
            objTable.Col = 9: .ReverseFg = objTable.Value
        End With
        objBarConfig.BarData(i) = mudtBarData(i)
    Next
    
End Sub

Private Sub ClearRtn()
    cboClientKind.ListIndex = 0
    cboBarKind.ListIndex = 0
    cboProject.ListIndex = -1
    cboPort.ListIndex = -1
    cboClientPort.ListIndex = -1
    txtWidth.Text = ""
    txtLength.Text = ""
    txtTotLen.Text = ""
    txtGapLen.Text = ""
    chkBarcode.Value = 0
    txtBarPosX.Text = ""
    txtBarPosY.Text = ""
    txtBarLength.Text = ""
    txtBarHeight.Text = ""
    cboStyle.ListIndex = -1
    chkAccCheck.Value = 0
    chkStat.Value = 0
    chkErFg.Value = 0
    txtStatPosX.Text = ""
    txtStatPosY.Text = ""
    txtLineLength.Text = ""
    txtLineWidth.Text = ""
    txtERcode.Text = ""
    cboReverse.ListIndex = -1
    tblDetails.Row = 1: tblDetails.Row2 = tblDetails.MaxRows
    tblDetails.Col = 2: tblDetails.Col2 = tblDetails.MaxCols
    tblDetails.BlockMode = True
    tblDetails.Action = ActionClearText
    tblDetails.BlockMode = False
    tblBldDetails.Row = 1: tblBldDetails.Row2 = tblBldDetails.MaxRows
    tblBldDetails.Col = 2: tblBldDetails.Col2 = tblBldDetails.MaxCols
    tblBldDetails.BlockMode = True
    tblBldDetails.Action = ActionClearText
    tblBldDetails.BlockMode = False
End Sub

Private Sub Form_Load()
    
    Me.Show
    DoEvents
    
    Call medAlwaysOn(frmBarConfig, 1)
    strFontList = "0.DF" & vbTab & "1.SMALL" & vbTab & "2.MIDDLE" & vbTab & "3.LARGE" & vbTab & "4.XLARGE" & vbTab & "5.XXLARGE" & vbTab & "6.Max"
    If mvarProjectId <> "" Then cboProject.ListIndex = medComboFind(cboProject, mvarProjectId)

End Sub

Private Sub tblBldDetails_Click(ByVal Col As Long, ByVal Row As Long)
   
    If Col = 6 Or Col = 7 Then
        With tblBldDetails
            .Col = Col: .Row = Row
            .CellType = CellTypeComboBox
            .TypeComboBoxList = strFontList
            .TypeComboBoxEditable = False
            SendKeys "{Enter}"
        End With
    End If
    
End Sub

Private Sub tblBldDetails_LeaveCell(ByVal Col As Long, ByVal Row As Long, _
                                   ByVal NewCol As Long, ByVal NewRow As Long, _
                                   Cancel As Boolean)
    
    Dim strValue As String
    
    If Col = 6 Or Col = 7 Then
        With tblBldDetails
            .Col = Col: .Row = Row
            strValue = .Text
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignCenter
            .Text = strValue
        End With
    End If
    
    If NewCol = 6 Or NewCol = 7 Then Call tblDetails_Click(NewCol, NewRow)

End Sub

Private Sub tblDetails_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col = 6 Or Col = 7 Then
        With tblDetails
            .Col = Col: .Row = Row
            .CellType = CellTypeComboBox
            .TypeComboBoxList = strFontList
            .TypeComboBoxEditable = False
            SendKeys "{Enter}"
        End With
    End If
    
End Sub

Private Sub tblDetails_LeaveCell(ByVal Col As Long, ByVal Row As Long, _
                                 ByVal NewCol As Long, ByVal NewRow As Long, _
                                 Cancel As Boolean)
    
    Dim strValue As String
    
    If Col = 6 Or Col = 7 Then
        With tblDetails
            .Col = Col: .Row = Row
            strValue = .Text
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignCenter
            .Text = strValue
        End With
    End If
    
    If NewCol = 6 Or NewCol = 7 Then Call tblDetails_Click(NewCol, NewRow)

End Sub

Private Sub txtLength_Change()
    txtGapLen.Text = Val(txtTotLen.Text) - Val(txtLength.Text)
End Sub

Private Sub txtTotLen_Change()
    txtGapLen.Text = Val(txtTotLen.Text) - Val(txtLength.Text)
End Sub
