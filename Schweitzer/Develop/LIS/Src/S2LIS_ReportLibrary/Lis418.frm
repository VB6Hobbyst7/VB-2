VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm418BatchReport 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Lis418.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   10950
   Begin VB.PictureBox picESign 
      Height          =   500
      Left            =   5955
      ScaleHeight     =   435
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   1485
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "출   력 (&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "임상병리 결과지 출력 조건"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   2145
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "결과지 출력 예정 리스트"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   14
      Top             =   270
      Width           =   10740
      Begin MSComCtl2.DTPicker dtpVfyToDt 
         Height          =   375
         Left            =   3390
         TabIndex        =   48
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   63635459
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpVfyDt 
         Height          =   375
         Left            =   1905
         TabIndex        =   18
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   63635459
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   870
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "보고일자"
         Appearance      =   0
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "종검"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   2
         Left            =   855
         TabIndex        =   43
         Top             =   405
         Width           =   885
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   4755
         ScaleHeight     =   375
         ScaleWidth      =   5805
         TabIndex        =   19
         Top             =   180
         Width           =   5865
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00F7F7F7&
            Caption         =   "개별 재출력"
            Height          =   375
            Index           =   2
            Left            =   4350
            Style           =   1  '그래픽
            TabIndex        =   31
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FFF4FD&
            Caption         =   "일괄 재출력"
            Height          =   375
            Index           =   1
            Left            =   2880
            Style           =   1  '그래픽
            TabIndex        =   30
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FEF5F3&
            Caption         =   "결과보고"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   0
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00EEFFFE&
            Caption         =   "회진용"
            Height          =   375
            Index           =   3
            Left            =   1485
            Style           =   1  '그래픽
            TabIndex        =   20
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   17
         Top             =   165
         Width           =   765
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   405
         Width           =   765
      End
      Begin VB.Label lblDash 
         Caption         =   "-"
         Height          =   225
         Left            =   3240
         TabIndex        =   49
         Top             =   300
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   5970
      Left            =   75
      ScaleHeight     =   5910
      ScaleWidth      =   10680
      TabIndex        =   3
      Top             =   2445
      Width           =   10740
      Begin FPSpread.vaSpread tblOrder 
         Height          =   5880
         Left            =   15
         TabIndex        =   12
         Top             =   0
         Width           =   10665
         _Version        =   196608
         _ExtentX        =   18812
         _ExtentY        =   10372
         _StockProps     =   64
         BackColorStyle  =   3
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   20
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis418.frx":06EA
         Appearance      =   1
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5910
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   10680
         _Version        =   196608
         _ExtentX        =   18838
         _ExtentY        =   10425
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
         GrayAreaBackColor=   14411494
         GridColor       =   14013909
         GridShowVert    =   0   'False
         MaxCols         =   46
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "Lis418.frx":12C4
         TextTip         =   4
      End
      Begin FPSpread.vaSpread tblList 
         Height          =   5910
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   10680
         _Version        =   196608
         _ExtentX        =   18838
         _ExtentY        =   10425
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   9
         MaxRows         =   50
         OperationMode   =   1
         ShadowColor     =   15857140
         SpreadDesigner  =   "Lis418.frx":23EF
         UserResize      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   1
      Top             =   915
      Width           =   10740
      Begin VB.OptionButton optTestDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "미생물"
         Height          =   255
         Index           =   2
         Left            =   9495
         TabIndex        =   42
         Top             =   165
         Width           =   855
      End
      Begin VB.OptionButton optTestDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기타"
         Height          =   255
         Index           =   1
         Left            =   8775
         TabIndex        =   41
         Top             =   165
         Width           =   705
      End
      Begin VB.OptionButton optTestDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일반"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   40
         Top             =   165
         Width           =   705
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "모두"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C56152&
         Height          =   255
         Left            =   7170
         TabIndex        =   39
         Top             =   165
         Width           =   765
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00FEF5F3&
         Caption         =   "미리보기(&V)"
         Height          =   510
         Left            =   9240
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   570
         Width           =   1320
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   7920
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   570
         Width           =   1320
      End
      Begin VB.Frame fraLabNo 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   750
         Left            =   210
         TabIndex        =   5
         Top             =   285
         Visible         =   0   'False
         Width           =   6420
         Begin VB.TextBox txtPtId 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1185
            TabIndex        =   9
            Text            =   "S00"
            Top             =   165
            Width           =   1275
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   15
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   165
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "환자 ID"
            Appearance      =   0
         End
         Begin VB.Label lblWard 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "61W-111"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4470
            TabIndex        =   36
            Top             =   270
            Width           =   690
         End
         Begin VB.Label lblSexAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "남/30"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3390
            TabIndex        =   34
            Top             =   255
            Width           =   450
         End
         Begin VB.Label lblPtNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "환자명1"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   2505
            TabIndex        =   33
            Top             =   255
            Width           =   630
         End
      End
      Begin VB.Frame fraSetWard 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   915
         Left            =   195
         TabIndex        =   21
         Top             =   165
         Width           =   6060
         Begin VB.CheckBox chkAllDoct 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2715
            TabIndex        =   27
            Top             =   570
            Width           =   705
         End
         Begin VB.TextBox txtDoctId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1230
            TabIndex        =   26
            Top             =   510
            Width           =   1065
         End
         Begin VB.CommandButton cmdDoctList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2310
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   25
            Top             =   495
            Width           =   315
         End
         Begin VB.TextBox txtWardId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1230
            TabIndex        =   24
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmdWardList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2310
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   23
            Top             =   120
            Width           =   315
         End
         Begin VB.CheckBox chkAllWard 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체병동/진료과"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2715
            TabIndex        =   22
            Top             =   195
            Width           =   1725
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   30
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "병동/진료과"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   6
            Left            =   30
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   510
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "주치의"
            Appearance      =   0
         End
         Begin VB.Label lblWardNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "WardNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4455
            TabIndex        =   32
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblDoctNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "DoctNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3465
            TabIndex        =   28
            Top             =   600
            Width           =   675
         End
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   " ☞ 출력대상자 리스트에서 선택하시면 출력 시 제외됩니다."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   8535
      Width           =   5955
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   11
      Top             =   8730
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 보고서 출력예정 건수 :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   8775
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   570
      Index           =   0
      Left            =   180
      Shape           =   4  '둥근 사각형
      Top             =   8430
      Width           =   6255
   End
End
Attribute VB_Name = "frm418BatchReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSql As New clsLISSqlReport

Private strStartDate As String
Private strEndDate As String
Private MsgFg As Boolean
Private PtFg As Boolean
Private ClearFg As Boolean

Dim blnLoadChk  As Boolean

Public Event FormClose()

Private Sub chkAll_Click()
    If chkAll.Value = 1 Then
        optTestDiv(0).Value = False
        optTestDiv(1).Value = False
        optTestDiv(2).Value = False
        optTestDiv(0).Enabled = False
        optTestDiv(1).Enabled = False
        optTestDiv(2).Enabled = False
    Else
        optTestDiv(0).Value = True
        optTestDiv(1).Value = False
        optTestDiv(2).Value = False
        optTestDiv(0).Enabled = True
        optTestDiv(1).Enabled = True
        optTestDiv(2).Enabled = True
    End If
End Sub

Private Sub chkAllDoct_Click()
    If optPrint(3).Value = True Then
        chkAllWard.Enabled = Choose(chkAllDoct.Value + 1, True, False)
    End If
    lblDoctNm.Caption = ""
    txtDoctId.Text = Choose(chkAllDoct.Value + 1, "", CS_AllCaption)
    txtDoctId.Enabled = Choose(chkAllDoct.Value + 1, True, False)
    cmdDoctList.Enabled = Choose(chkAllDoct.Value + 1, True, False)
End Sub

Private Sub chkAllWard_Click()
    If optPrint(3).Value = True Then
        chkAllDoct.Enabled = Choose(chkAllWard.Value + 1, True, False)
    End If
    lblWardNm.Caption = ""
    txtWardId.Text = Choose(chkAllWard.Value + 1, "", CS_AllCaption)
    txtWardId.Enabled = Choose(chkAllWard.Value + 1, True, False)
    cmdWardList.Enabled = Choose(chkAllWard.Value + 1, True, False)
End Sub


Private Sub cmdClear_Click()
    TxtClear
End Sub

Private Sub cmdDoctList_Click()

'% 주치의 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
'    Dim objDoct As New clsBasisData

    With objMyList
        .Connection = DBConn
        .FormCaption = "주치의리스트"
        .ColumnHeaderText = "의사ID;의사명"
        .Tag = "DoctID"
        Me.ScaleMode = 1
'        Call .ListPop(GetDoctListSQL, 3950, 6300)
        Call .LoadPopUp(GetSQLDoctList) ', 3950, 6300)

        txtDoctId.Text = medGetP(.SelectedString, 1, ";")
        lblDoctNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objDoct = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdExit_Click()
    Set objSql = Nothing
    Unload Me
    
    RaiseEvent FormClose
End Sub

Private Sub cmdPreview_Click()
    
    Dim i As Long
    Dim strPtId    As String
    Dim strPtNm    As String
    Dim strVfyDt   As String
    Dim strTestDiv As String
    
    If cmdPreview.Tag = "1" Then
        cmdPreview.Caption = "미리보기(&V)"
        cmdPreview.Tag = ""
        tblOrdSheet.Visible = False
        tblOrdSheet.ZOrder 1
    Else
        If tblOrder.MaxRows = 0 Then Exit Sub
        
        Dim objProgress As New clsProgress
        
        objProgress.Container = MainFrm.stsbar
        objProgress.Message = "자료를 읽고 있습니다..."
        objProgress.Max = tblOrder.MaxRows
'        objProgress.Caption = "처리중입니다."
'        objProgress.Mode = 0
'        objprogress.message = "자료를 읽고 있습니다."
'        objProgress.Max = tblOrder.MaxRows
'        objProgress.Min = 0
'        objProgress.Value = 0
'        objProgress.Visible = True
        
        tblOrdSheet.MaxRows = 0
        
        For i = 1 To tblOrder.MaxRows
            objProgress.Value = i
            tblOrder.Row = i
            tblOrder.Col = 1
            If tblOrder.Value = 1 Then GoTo Skip
            tblOrder.Col = 4:  strPtId = tblOrder.Value
            tblOrder.Col = 5:  strPtNm = tblOrder.Value
            tblOrder.Col = 16: strTestDiv = tblOrder.Value
            strVfyDt = Format(dtpVfyDt.Value, CS_DateDbFormat)
            objProgress.Message = strPtNm & "환자의 결과내역을 읽고 있습니다."
            DoEvents
            Call DisplayOrders(strPtId, strPtNm, strVfyDt, strTestDiv)
Skip:
        Next
        cmdPreview.Caption = "닫기(&B)"
        cmdPreview.Tag = "1"
        tblOrdSheet.Visible = True
        tblOrdSheet.ZOrder 0
    End If
End Sub

Private Sub cmdWardList_Click()
'% 병동코드 리스트를 팝업한다.

    Dim objMyList As New clsPopUpList
    Dim strCaption As String
    Dim strHead As String
    
    If optBussDiv(0).Value Then
        strCaption = "진료과 조회"
        strHead = "부서코드;부서명"
    Else
        strCaption = "병동 조회"
        strHead = "병동코드;병동명"
    End If
    
'    Dim objDept As clsBasisData
    
'    Set objDept = New clsBasisData
    
    With objMyList
        .Connection = DBConn
        .FormCaption = strCaption
        .ColumnHeaderText = strHead
        .Tag = "WardID"
        Me.ScaleMode = 1
        
        If optBussDiv(0).Value Then
'            Call .ListPop(, 3950, 6300, ObjLISComCode.DeptCd)
            Call .LoadPopUp(GetSQLDeptList) ', 3950, 6300)
        Else
'            Call .ListPop(, 3950, 6300, ObjLISComCode.WardID)
            Call .LoadPopUp(GetSQLWardList) ', 3950, 6300)
        End If
        
        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        lblWardNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objDept = Nothing
    Set objMyList = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub

Private Sub cmdPrint_Click()
    
    Dim strRstEntryType As String
    Dim strPtId         As String
    Dim strTestDiv      As String
    Dim strTable        As String
    Dim strSQL          As String
    Dim strImgPath      As String
    Dim strVfyDt        As String
    Dim strVfyTm        As String
    Dim i               As Long
    Dim j               As Long
    Dim objProgress     As jProgressBar.clsProgress
    Dim objProgress1    As jProgressBar.clsProgress
    Dim objReport       As clsBatchReport
    Dim strLastDt       As String
    Dim strLastTm       As String
    Dim strPrtDt        As String
    Dim strPrtTm        As String
    Dim lngErrCount     As Long
    
    Dim lngFileNo As Long
    
    lngFileNo = FreeFile
    
    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtWardId.Text) = "" Then
        MsgBox "결과지를 출력할 병동을 선택하십시오.", vbInformation, "병동선택"
        txtWardId.SetFocus
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtDoctId.Text) = "" Then
        MsgBox "주치의를 선택하십시오.", vbInformation, "주치의선택"
        txtDoctId.SetFocus
        Exit Sub
    End If
    
    If lblCnt.Caption = 0 Then
        MsgBox "출력할 대상 리스트가 없습니다.", vbInformation, "결과 출력"
        Exit Sub
    End If
    
    If optPrint(3).Value = True Then
        Call ReportPrint
        Exit Sub
    End If
    
    lngErrCount = 0
    
    MouseRunning
    
    Set objProgress = New jProgressBar.clsProgress
    
    With objProgress
        .Container = Me
        .Max = tblOrder.MaxRows
        .Left = lblPrgBar.Left + 3
        .Top = lblPrgBar.Top + 3
        .Width = lblPrgBar.Width - 10
        .Height = lblPrgBar.Height - 10
        
'        .SetMyForm Me
'        .Choice = True
'        .Max = tblOrder.MaxRows
'        .Min = 0
'        .Value = 0
'        .XPos = lblPrgBar.Left + 3
'        .YPos = lblPrgBar.Top + 3
'        .XWidth = lblPrgBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblPrgBar.Height - 10 ' 260
        DoEvents
    End With
    
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        If optBussDiv(0).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_외래.log" For Append As lngFileNo
        ElseIf optBussDiv(1).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_병동.log" For Append As lngFileNo
        Else
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_종검.log" For Append As lngFileNo
        End If
    End If
    
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    With tblOrder
        
        For i = 1 To .MaxRows
            
On Error GoTo Err_Trap1

            objProgress.Value = i
            
            .Row = i
            
            .Col = 1
            If .Value = 0 Then
                
                .TopRow = i
                
                .Col = 5    '환자명
                objProgress.Message = .Value & " 환자의 결과지를 출력하고 있습니다... ( " & i & " / " & .MaxRows & " )"

                .Col = 4    '환자ID
                strPtId = .Value
                
                .Col = 15   '전자서명 Path
                strImgPath = .Value

                .Col = 16   '보고서 종류
                strTestDiv = .Value
                
                picESign.Picture = LoadPicture(strImgPath)

                Set objReport = New clsBatchReport

                'Dictionary에 담기..레포트 출력
                .Col = 2:
                
'                Set objDept = Nothing
'                Set objDept = New clsBasisData
                strDept = GetWardNm(medGetP(.Value, 1, "-"))
'                Set objDept = Nothing
                
                If strDept <> "" Then
                    objReport.Ward = strDept
                    If objReport.Ward <> "" Then
                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    Else
                        objReport.Ward = Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    End If
                End If
                
'                If ObjLISComCode.WardID.Exists(medgetp(.Value, 1, "-")) = True Then
'                    ObjLISComCode.WardID.KeyChange (medgetp(.Value, 1, "-"))
'                    objReport.Ward = ObjLISComCode.WardID.Fields("wardnm")
'
'                    If objReport.Ward <> "" Then
'                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    Else
'                        objReport.Ward = Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    End If
'                End If
                
                .Col = 3:  objReport.Doct = .Value
                .Col = 4:  objReport.ptid = .Value
                .Col = 5:  objReport.PtNm = .Value
                .Col = 6:  objReport.PtSex = medGetP(.Value, 1, "/")
                           objReport.PtAge = medGetP(.Value, 2, "/")
                .Col = 10: strVfyDt = .Value
                .Col = 11: strVfyTm = .Value
                
                objReport.VfyDt = strVfyDt & " " & strVfyTm
                
                '.Col = 11: objReport.VfyDt = objReport.VfyDt & " " & .Value
                .Col = 12: objReport.VfyNM = .Value
                .Col = 13: objReport.MdfDt = .Value         '수정일
                .Col = 17: objReport.ICD = .Value
                
                
                '병동에서 출력할때만 레포트제목에 재발행/회진용 표기
                If gUsingInWardMenu Then
                    objReport.Rouding = optPrint(3).Value       '회진레포트 여부
                    objReport.Reprint = optPrint(2).Value       '재발행 여부
                    objReport.BatchReprint = True
                Else
                    objReport.Rouding = optPrint(3).Value       '회진레포트 여부
                    objReport.Reprint = optPrint(2).Value       '재발행 여부
                    objReport.BatchReprint = optPrint(1).Value
                End If
                objReport.Special = IIf(strTestDiv = enTestDiv.TST_SpeTest, True, False)
                
                .Col = 18:
                
'                Set objDept = Nothing
'                Set objDept = New clsBasisData
                strDept = GetDeptNm(.Value)
'                Set objDept = Nothing
                
                If strDept <> "" Then
                    objReport.Dept = .Value
                    objReport.DeptNm = strDept
                Else
                    objReport.Dept = .Value
                End If
                
'                If ObjLISComCode.DeptCd.Exists(.Value) Then
'                    Call ObjLISComCode.DeptCd.KeyChange(.Value)
'                    objReport.Dept = .Value
'                    objReport.DeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                Else
'                    objReport.Dept = .Value
'                End If

                
                If optPrint(0).Value And (Not gUsingInWardMenu) Then
                    Print #lngFileNo, "( " & i & " / " & .MaxRows & " )  " & Now & "   " & strPtId & "," & objReport.PtNm & "," & objReport.DeptNm & "," & objReport.Ward
                End If
                
                If optPrint(2).Value = True Then
                    Call objReport.ReportForOnePatient(strPtId, strStartDate, Format(dtpVfyToDt.Value, CS_DateDbFormat), _
                                                       strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
                Else
                    Call objReport.ReportForOnePatient(strPtId, strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), _
                                                       strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
                End If
            End If
'            objProgress.value = objProgress.value + 1
    
            '결과보고인 경우 & 병동메뉴가 아닌경우...
    
        Next
    End With

    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        
        strPrtDt = Format(GetSystemDate, CS_DateDbFormat)
        strPrtTm = Format(GetSystemDate, CS_TimeDbFormat)
        
        strSQL = objSql.UpdatePrtDt(strPrtDt, strPrtTm, ObjSysInfo.EmpId)
On Error GoTo Err_Trap
            
        DBConn.BeginTrans
        DBConn.Execute strSQL
        DBConn.CommitTrans
    End If
    
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Close #lngFileNo
    End If

    MouseDefault
    
    Set objProgress = Nothing
    Set objProgress1 = Nothing
    
    If lngErrCount > 0 Then
        For i = tblOrder.DataRowCnt To 1 Step -1
            tblOrder.Row = i
            tblOrder.Col = 20
            If tblOrder.Value = "0" Then
                tblOrder.Action = ActionDeleteRow
            End If
        Next
        MsgBox "다음 환자들의 결과지 출력 중 오류가 발생했습니다. 다시 출력하십시오.", vbExclamation, "오류"
    Else
        cmdClear_Click
    End If
    
    Exit Sub
    
Err_Trap:
    DBConn.RollbackTrans
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "DB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap
    Resume Next

Err_Trap1:
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "VB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap1
    Resume Next
    
End Sub

Private Sub ReportPrint()
    Dim strDeptNm As String
    Dim strDoctNm As String
    Dim strFont   As String
     
    strDeptNm = IIf(txtWardId.Text = "", "전체", lblWardNm.Caption)
    strDoctNm = IIf(txtDoctId.Text = "", "전체", lblDoctNm.Caption)
    
     With tblList
        If Printer.PaperSize = vbPRPSA4 Then
            .PrintMarginLeft = 500
            .PrintMarginRight = 800
        Else
            .PrintMarginTop = 300
            .PrintMarginBottom = 100
            .PrintMarginLeft = 200
            .PrintMarginRight = 100
        End If
        .PrintJobName = "회진용레포트 출력"
    
        .PrintAbortMsg = "회진용레포트 출력중입니다. "
    
        .PrintColor = False
        .PrintFirstPageNumber = 1
        strFont = "/fn""굴림체""/fz""11"""
        .PrintHeader = strFont & "/n/n/l/fb1 " & "☞ 임상병리 회진용 결과 보고서 : " & Format(dtpVfyDt.Value, CS_DateLongFormat) & " /n/l/fb1   병동:" & strDeptNm & _
        ",주치의:" & strDoctNm & "(" & lblCnt.Caption & "명) /c/fb1/n"
        .PrintFooter = " /l " & String(103, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1" & " /r 임상병리 회진용 보고서 "
        .PrintShadows = False
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = True
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll
    
        .Action = ActionPrint
    
        .GridSolid = True
    End With

End Sub


Private Sub cmdQuery_Click()

    Dim objReport   As New clsBatchReport
    Dim objESign    As clsLISElectronSign
    Dim objProgress As clsProgress
    Dim Rs          As New Recordset
    Dim strWA       As String
    Dim strTable    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strReferral As String
    Dim strSex      As String
    Dim strStsCd    As String
    Dim strMsg      As String
    Dim i           As Long
    Dim strEmpId    As String
    Dim strBussDiv  As String
    Dim strChkLoad  As String
    Dim strTestDiv  As String
    Dim strKey      As String
    Dim strDOB      As String
    Dim strSQL      As String
    Dim strWard     As String
    Dim strDoct     As String
    Dim strPtNm     As String
    
    tblOrder.MaxRows = 0
    lblCnt.Caption = 0
    Me.MousePointer = 11
    strBussDiv = IIf(optBussDiv(0).Value, enBussDiv.BussDiv_OutPatient, enBussDiv.BussDiv_InPatient)
    
    'strStartDate = objReport.GetLastReportDt
    If optPrint(2).Value = True Or optPrint(3).Value = True Then
        strStartDate = Format(dtpVfyDt.Value, CS_DateDbFormat)
    Else
        strStartDate = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    End If
    
    '-- 2007.06.28 osw
    If optPrint(2).Value = True Then
        strEndDate = Format(dtpVfyToDt.Value, CS_DateDbFormat)
    End If
    
    '프로그래스바 생성..
    Set objProgress = New clsProgress
    objProgress.Container = MainFrm.stsbar
    objProgress.Message = "자료를 읽고 있습니다..."
    objProgress.Max = 100
'    objProgress.Caption = "처리중입니다."
'    objProgress.Mode = 0
'    objprogress.message = "자료를 읽고 있습니다."
'    objProgress.Max = 100
'    objProgress.Min = 0
'    objProgress.Value = 0
'    objProgress.Visible = True
    
    If txtWardId.Text <> CS_AllCaption Then strWard = txtWardId.Text
    If txtDoctId.Text <> CS_AllCaption Then strDoct = txtDoctId.Text

    '개별재출력
    If optPrint(2).Value = True Then
        '-- 2007.06.28 osw
'        Rs.Open objSql.GetAccLAbNoLIS201(txtPtId.Text, Format(dtpVfyDt.Value, CS_DateDbFormat)), DBConn
        Rs.Open objSql.GetAccLAbNoLIS201_New(txtPtId.Text, Format(dtpVfyDt.Value, CS_DateDbFormat), Format(dtpVfyToDt.Value, CS_DateDbFormat)), DBConn
        tblOrder.ZOrder 0
    
    '일괄 정기출력
    ElseIf optPrint(0).Value = True Then
        strTestDiv = IIf(chkAll.Value = 1, "3", IIf(optTestDiv(0).Value, "0", IIf(optTestDiv(1).Value, "1", "2")))
        
        '결과지 미출력 대상 RPTFG 박아주기
        If P_NoResultReport <> "" Then
            Call objReport.SetNoReportRptUpdate(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat))
        End If
        
        Rs.Open objSql.LABReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "", strTestDiv, strWard, strDoct), DBConn
        
        tblOrder.ZOrder 0
    
    '일괄재출력
    ElseIf optPrint(1).Value = True Then
        strTestDiv = IIf(chkAll.Value = 1, "3", IIf(optTestDiv(0).Value, "0", IIf(optTestDiv(1).Value, "1", "2")))
        Rs.Open objSql.LABReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "Y", strTestDiv, strWard, strDoct), DBConn
        tblOrder.ZOrder 0
    '병동회진용
    Else
        tblList.ZOrder 0
        strTestDiv = IIf(chkAll.Value = 1, "3", IIf(optTestDiv(0).Value, "0", IIf(optTestDiv(1).Value, "1", "2")))
        Call DisplayReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), "2" & COL_DIV & strTestDiv, strBussDiv, _
                                    txtWardId.Text, txtDoctId.Text)
        Me.MousePointer = 0
        Set objReport = Nothing
        Set Rs = Nothing
        Exit Sub
    End If
        
    If Rs.EOF Then
        Set objProgress = Nothing
        MsgBox "해당 데이타가 없습니다.", vbInformation, "결과지 출력"
        GoTo Nodata
    End If
    
    strKey = ""
    With tblOrder
        If Rs.RecordCount > 0 Then
            '프로그래스바 생성..
            objProgress.Max = Rs.RecordCount
            objProgress.Min = 0
            objProgress.Value = 0
            .ReDraw = False
            i = 1
            
'            Dim objEmp As clsBasisData
            Dim strEmp As String
            
            Do Until Rs.EOF = True
                If strKey = "" & Rs.Fields("deptcd").Value & _
                                 Rs.Fields("ptid").Value & _
                                 Rs.Fields("testdiv").Value Then
                    '환자/진료과/보고서종류가 같은 경우엔 수정여부와 수정일만 보여주기...
                    If "" & Rs.Fields("stscd").Value = enStsCd.StsCd_LIS_Modify Then
                        .Col = 9
                        .Value = "수정"
                    End If
                    If Trim("" & Rs.Fields("mfydt").Value) <> "" Then
                        .Col = 13
                        .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                    End If
                    GoTo Skip
                End If
                    
                .MaxRows = i
                .Row = i
                
                
                .Col = 2:   .Value = Rs.Fields("location").Value & ""
                .Col = 19:  .Value = Rs.Fields("location").Value & ""
                .Col = 18:  .Value = Rs.Fields("deptcd").Value & ""
                
                If optBussDiv(1).Value Then
                    If Rs.Fields("location").Value & "" <> "" Then
                        .Col = 2:   .Value = Rs.Fields("location").Value & "" '& "-" & Rs.Fields("hosilid").Value & ""
                        .Col = 19:  .Value = Rs.Fields("location").Value & ""
                    Else
                        .Col = 2:   .Value = Rs.Fields("deptcd").Value & ""
                    End If
                ElseIf optBussDiv(0).Value Then
                    .Col = 2:       .Value = Rs.Fields("deptcd").Value & ""
                    .Col = 19:      .Value = Rs.Fields("location").Value & ""
                End If
                
                If optPrint(2).Value Then
                    If lblWard.Caption <> "" Then
                        .Col = 2:   .Value = lblWard.Caption
                        .Col = 19:  .Value = lblWard.Caption
                    End If
                End If
                
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmp = Nothing
                
'                .Col = 3: .Value = GetEmpName("" & rs.Fields("majdoct").Value)
                .Col = 3: .Value = strEmp
                
                .Col = 4: .Value = "" & Rs.Fields("ptid").Value
                
'                Call GetPatientInfo(rs.Fields("ptid").Value & "", strPtNm, strSex, strDOB)
                
                Dim objPt As clsPatient
                
                Set objPt = Nothing
                Set objPt = New clsPatient
                
                Call objPt.GETPatient(Rs.Fields("ptid").Value & "")
                
                .Col = 5: .Value = objPt.PtNm ' strPtNm
                

                If IsNumeric(objPt.Sex) Then 'strSex
                    strSex = IIf(Val(objPt.Sex) Mod 2 = 1, "남", "여")
                Else
                    strSex = IIf(objPt.Sex = "M", "남", "여")
                End If
                
                
                .Col = 6: .Value = strSex
                           strDOB = Mid(objPt.Dob, 1, 6) 'strDOB
                           If Len(strDOB) = 6 Then strDOB = strDOB & "01"
                           If IsDate(Format(strDOB, CS_DateMask)) Then
                                .Value = .Value & "/" & DateDiff("yyyy", Format(strDOB, CS_DateMask), Now)
                           Else
                                .Value = .Value & "/미확"
                           End If
                Set objPt = Nothing
                
                .Col = 16
                .Value = "" & Rs.Fields("testdiv").Value
                
                .Col = 7
                Select Case "" & Rs.Fields("testdiv").Value
                Case enTestDiv.TST_RouTest
                    .Value = "일반"
                Case enTestDiv.TST_SpeTest
                    .Value = "기타"
                Case enTestDiv.TST_MicTest
                    .Value = "미생물"
                End Select

                .Col = 8: .Value = 1
                .Col = 9
                Select Case "" & Rs.Fields("stscd").Value
                Case enStsCd.StsCd_LIS_MidRst
                    .Value = "중간"
                Case enStsCd.StsCd_LIS_FinRst
                    .Value = "최종"
                Case enStsCd.StsCd_LIS_Modify
                    .Value = "수정"
                End Select

                .Col = 10: .Value = Format(Mid("" & Rs.Fields("vfydt").Value, 3), CS_DateShortMask)
                .Col = 11: .Value = Format(Mid("" & Rs.Fields("vfytm").Value, 1, 4), CS_TimeShortMask)

                strEmpId = "" & Rs.Fields("vfyid").Value
                .Col = 14: .Value = strEmpId
                
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("vfyid").Value & "")
'                Set objEmp = Nothing
                
'                .Col = 12: .Value = GetEmpName(strEmpId)
                .Col = 12: .Value = strEmp
                
                If .Value <> "" Then
                    Set objESign = New clsLISElectronSign
                    If objESign.LoadElectronSign(strEmpId, InstallDir & "LIS\") = True Then
                        If objESign.ElectronSignPrintOk = True Then
                            .ForeColor = vbBlue
                            .Col = 15: .Value = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
                        Else
                            .ForeColor = vbBlack
                        End If
                    End If
                    Set objESign = Nothing
                End If

                .Col = 13: .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                '임상진단....
                Dim objDisease  As New clsDisease
                
                objDisease.ptid = Rs.Fields("ptid").Value
                
                .Col = 17: .Value = objDisease.Disease
                
                Set objDisease = Nothing
                
                strKey = "" & Rs.Fields("deptcd").Value & _
                              Rs.Fields("ptid").Value & _
                              Rs.Fields("testdiv").Value
                              
                i = i + 1
                objProgress.Value = objProgress.Value + 1
Skip:
                Rs.MoveNext
            Loop
            Set objProgress = Nothing
            .ReDraw = True
            lblCnt.Caption = .MaxRows
        Else
            If optPrint(0).Value = True Then
                strMsg = "결과보고"
            ElseIf optPrint(1).Value = True Then
                strMsg = "일괄재출력"
            ElseIf optPrint(2).Value = True Then
                strMsg = "개별재출력"
            Else
                strMsg = "회진용 결과지 출력"
            End If
            MsgBox strMsg & " 내역이 없습니다.", vbCritical, "결과 출력"
            medClearTable tblOrder
            tblOrder.MaxRows = 0
            lblCnt.Caption = 0
        End If
        .Refresh
        .ZOrder 0
    End With

Nodata:
    Me.MousePointer = 0
    Set Rs = Nothing
    Set objProgress = Nothing

End Sub

Private Sub DisplayReportList(ByVal pStartDate As String, ByVal pVfyDt As String, _
                              ByVal pChkLoad As String, ByVal pBussDiv As String, ByVal pWardId As String, _
                              ByVal pDoctId As String)
    
    Dim Rs As New Recordset
    Dim rs1 As New Recordset
    Dim objProgress As clsProgress
    Dim strPtId As String
    Dim ii As Long
    Dim jj As Long
    Dim kk As Long
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
        
    Rs.Open objSql.GetLisReportList(pStartDate, pVfyDt, pChkLoad, pBussDiv, pWardId, pDoctId), DBConn

    If Rs.EOF Then GoTo Nodata
    
    '프로그래스바 생성..
    Set objProgress = New clsProgress
    objProgress.Container = MainFrm.stsbar
    objProgress.Message = "자료를 읽고 있습니다..."
    objProgress.Max = Rs.RecordCount
'    objProgress.Caption = "처리중입니다."
'    objProgress.Mode = 0
'    objprogress.message = "자료를 읽고 있습니다."
'    objProgress.Max = Rs.RecordCount
'    objProgress.Min = 0
'    objProgress.Value = 0
'    objProgress.Visible = True

    ii = 0
    jj = 2
    With tblList
        .ReDraw = False
        Rs.MoveFirst
        strPtId = ""
        Do Until Rs.EOF
                If strPtId <> Rs.Fields("ptid").Value & "" And strPtId <> "" Then ii = 0: jj = jj + 1
                strPtId = Rs.Fields("ptid").Value & ""
                .MaxCols = jj
                .Row = ii
                .Col = jj
                .ColWidth(jj) = 8
                
                If .Row = 0 Then
                    .Value = Rs.Fields("ptid").Value & "" & vbNewLine & Rs.Fields("ptnm").Value & "" & vbNewLine & Rs.Fields("wardid").Value & "" & "/" & Rs.Fields("hosilid").Value & ""
                End If
                
                For kk = 2 To .MaxRows
                    .Row = kk
                    .Col = 1
                    If Trim(.Value) = Rs.Fields("testcd").Value & "" Then
                        .Col = jj
                        If Trim(Rs.Fields("rsttype").Value & "") = "A" Then
                            Set rs1 = Nothing
                            Set rs1 = New Recordset
                            rs1.Open objSql.GetVfyTestCd(Rs.Fields("testcd").Value & "", Rs.Fields("rstcd").Value & ""), DBConn
                            If rs1.RecordCount > 0 Then
                                .Value = rs1.Fields("field1").Value & ""
                            End If
                            Set rs1 = Nothing
                        Else
                            .Value = Rs.Fields("rstcd").Value & ""
                        End If
                    End If
                Next kk
                
            ii = ii + 1
            objProgress.Value = objProgress.Value + 1
            Rs.MoveNext
        Loop
        .ReDraw = True
    End With
    lblCnt.Caption = jj - 1
Nodata:
    Set Rs = Nothing
    Set rs1 = Nothing
    Set objProgress = Nothing
End Sub


Private Sub Form_Load()
    lblWardNm.Caption = ""
    lblWard.Caption = ""
    
    optBussDiv(2).Visible = False
    
    If gUsingInWardMenu Then
        optPrint(1).Visible = False
        'optPrint(2).Visible = False
        'optBussDiv(0).Enabled = False
        'optBussDiv(1).Value = True
        chkAllWard.Value = 0
        chkAllWard.Visible = False
        chkAllDoct.Value = 1
    Else
        optPrint(1).Visible = True
        optPrint(2).Visible = True
        'optBussDiv(0).Enabled = True
        'optBussDiv(0).Value = True
        chkAllWard.Value = 1
        chkAllDoct.Value = 1
    End If
    
    
    
    optBussDiv(0).Enabled = True
    optBussDiv(0).Value = True

'    Me.Top = 0
'    Me.Height = frmReportTree.Height
'    Me.Width = medMain.Width - frmReportTree.Width - medMain.picComTool.Width - 200
'    Me.Left = frmReportTree.Width

    blnLoadChk = False
    TxtClear
End Sub

Private Sub TxtClear()
    '결과지 출력 조건
    dtpVfyDt.Value = GetSystemDate
    dtpVfyToDt.Value = GetSystemDate
    
    '결과지 출력예정리스트
    medClearTable tblOrder
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
    
    lblWard.Caption = ""
    tblOrder.MaxRows = 0
    tblOrdSheet.MaxRows = 0
    tblOrder.ZOrder 0
    chkAll.Value = 1
    txtWardId.Text = "(전체)"
    lblCnt.Caption = 0
    chkAllWard.Value = 1
    chkAllDoct.Value = 1
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    cmdPreview.Caption = "미리보기(&V)"
    cmdPreview.Tag = ""
    tblOrdSheet.Visible = False
    tblOrdSheet.ZOrder 1

End Sub


Private Sub optBussDiv_Click(Index As Integer)
    cmdClear_Click
End Sub

Private Sub optPrint_Click(Index As Integer)
    '-- 2007.06.28 osw
    lblDash.Visible = False
    dtpVfyToDt.Visible = False
    
    If optPrint(2).Value = True Then
        '-- 2007.06.28 osw
        lblDash.Visible = True
        dtpVfyToDt.Visible = True
        
        chkAll.Value = 1
        chkAll.Enabled = False
        fraLabNo.Visible = True
        fraSetWard.Visible = False
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        chkAll.Enabled = True
        fraLabNo.Visible = False
        fraSetWard.Visible = True
    End If
    
    If optPrint(0).Value = True Then
        If gUsingInWardMenu Then
            chkAllWard.Value = 0
            chkAllWard.Visible = False
        End If
    End If
    
    If optPrint(3).Value = True Then
        tblList.Visible = True
        tblList.ZOrder 0
        tblOrder.Visible = False
        
        optBussDiv(1).Value = True
        optBussDiv(0).Enabled = False
        chkAllDoct.Value = 0
        chkAllWard.Value = 0
        chkAllWard.Visible = True
'        chkAllDoct.Enabled = False
'        chkAllWard.Enabled = False
        
        Call GetTestlist
    Else
        optBussDiv(0).Enabled = True
        If gUsingInWardMenu Then
            chkAllDoct.Enabled = True
            chkAllDoct.Value = 1
        Else
            chkAllDoct.Enabled = True
            chkAllWard.Enabled = True
            chkAllDoct.Value = 1
            chkAllWard.Value = 1
        End If
        
        tblList.Visible = False
        tblOrder.Visible = True
        tblOrder.ZOrder 0
    End If

    dtpVfyDt.Value = GetSystemDate
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    '결과지 출력예정리스트
    medClearTable tblOrder

    lblCnt.Caption = 0

End Sub

Private Sub GetTestlist()
    Dim Rs As New Recordset
    Dim strTestNM As String
    Dim ii As Long
    Dim jj As Long
    
    
    Rs.Open objSql.GetTestReportList, DBConn
    If Rs.RecordCount > 0 Then
        ii = 0
        jj = 0
        strTestNM = ""
        With tblList
        
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = 1
            .BlockMode = True
            .AllowCellOverflow = False
            .BlockMode = False
            
            .ReDraw = False
            .MaxRows = Rs.RecordCount + 1
            .Row = ii: .Col = 0
            .Value = "검사명/등록번호" & vbNewLine & "환자명" & vbNewLine & "병동/병실"
            ii = 1
            Rs.MoveFirst
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = 0
                .RowHeight(ii) = 9.5
                If Rs.Fields("panelfg").Value & "" = "D" Then strTestNM = Rs.Fields("cdval1").Value & ""
                jj = Len(strTestNM)
                If strTestNM = Mid(Rs.Fields("cdval1").Value & "", 1, jj) And jj <> "0" And strTestNM <> Rs.Fields("cdval1").Value & "" Then
                    .Value = Space(4) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                Else
                    .Value = Space(1) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                End If
                .Col = 1
                .Value = Rs.Fields("cdval1").Value & "": .ForeColor = vbWhite
                Rs.MoveNext
            Loop
            .ReDraw = True
        End With
    End If
    Set Rs = Nothing
End Sub

'
Private Sub tblOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If MsgFg Then Exit Sub
    
    Dim lngButtonValue As Long
    Dim i As Long
    Dim strDept As String
    
    With tblOrder
        .Row = Row
        .Col = Col
        lngButtonValue = .Value
        If .Value = 1 Then
            lblCnt.Caption = Val(lblCnt.Caption) - 1
        Else
            lblCnt.Caption = Val(lblCnt.Caption) + 1
            Exit Sub
        End If
        
        .Col = 2
        strDept = medGetP(.Value, 1, "-")
        For i = 1 To tblOrder.DataRowCnt
            MsgFg = True
            .Row = i
            .Col = 2
            If strDept = medGetP(.Value, 1, "-") Then
                .Col = 1
                .Value = lngButtonValue
            End If
            MsgFg = False
        Next
    End With
End Sub

Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then txtPtId.SetFocus
End Sub

Private Sub tblOrder_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i As Long
    Static lngOnOff As Long
    
    If Row = 0 And Col = 1 Then
        lngOnOff = (lngOnOff + 1) Mod 2
        For i = 1 To tblOrder.MaxRows
            tblOrder.Row = i
            tblOrder.Col = 1
            tblOrder.Value = lngOnOff
        Next
    End If
    
End Sub

Private Sub tblOrdSheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    Dim tmpToolTip As String
    Dim strSQL As String
    Dim tmpColNm As String
   
    'If Not OrderFg Then Exit Sub
   
    tmpToolTip = vbCrLf
   
    With tblOrdSheet
        .Row = Row
       
        If Col = 6 Then
            .Col = 6
            If Len(.Value) > 20 Then
                MultiLine = 1
                TipWidth = 4000
                tmpToolTip = vbCrLf & Space(3) & .Value & Space(3) & vbCrLf
                TipText = tmpToolTip
                ShowTip = True
                Exit Sub
            End If
        End If
       
        .Col = 5:    If Trim(.Value) = "" Then Exit Sub
       
        'If chkToolTip.Value = 0 Then GoTo Skip
       
        .Col = 12:  tmpToolTip = tmpToolTip & "  처    방 : " & .Value       '처방일
        .Col = 13:  tmpToolTip = tmpToolTip & " ( # " & Format(.Value, "##") & " )" & vbCrLf        '처방번호
        .Col = 15:  tmpToolTip = tmpToolTip & "  채    혈 : " & .Value & vbCrLf       '채혈일시
        .Col = 17:  tmpToolTip = tmpToolTip & "  접    수 : " & .Value & vbCrLf      '접수일시
        .Col = 6:
                    If .Value <> "미확" Then
                        .Col = 25:   tmpToolTip = tmpToolTip & "  결과보고 : " & .Value & vbCrLf      '보고일시
                    End If
       .Col = 22:
                    If .Value <> "" Then
                        tmpToolTip = tmpToolTip & vbCrLf & "  최근결과 : [ " & .Value & " ] " '& vbCrLf        '최근결과
                        '.Col = 21:   tmpToolTip = tmpToolTip & "             " & .Value  '최근결과일시
                        .Col = 23
                        tmpToolTip = tmpToolTip & Mid(.Value, 1, 9) & vbCrLf '최근결과일시
                    End If
       
Skip:
     
Skip1:
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 500
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        'If chkToolTip.Value = 1 Then
            ShowTip = True
        'Else
        '    ShowTip = False
        'End If
       
    End With
   
End Sub

Private Sub txtDoctId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDoctId_LostFocus()
'    Dim objEmp As clsBasisData
    
    If Trim(txtDoctId.Text) = "" Then Exit Sub
    
'    Set objEmp = New clsBasisData
    
    
    lblDoctNm.Caption = GetEmpNm(txtDoctId.Text) 'GetEmpName(txtDoctId.Text)
    If lblDoctNm.Caption = "" Then
        txtDoctId.Text = ""
        lblDoctNm.Caption = ""
    End If
    
'    Set objEmp = Nothing
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtWardId_LostFocus()
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    If Trim(txtWardId.Text) = "" Then Exit Sub
    
'    Set objDept = New clsBasisData
    
    If optBussDiv(0).Value Then
        strDept = GetDeptNm(txtWardId.Text)
        
        If strDept <> "" Then
            lblWardNm.Caption = strDept
        Else
            txtWardId.Text = ""
            lblWardNm.Caption = ""
        End If
    Else
        strDept = GetWardNm(txtWardId.Text)
        
        If strDept <> "" Then
            lblWardNm.Caption = strDept
        Else
            txtWardId.Text = ""
            lblWardNm.Caption = ""
        End If
    End If
'    Set objDept = Nothing
'    If optBussDiv(0).Value = True Then
'        If ObjLISComCode.DeptCd.Exists(txtWardId.Text) = True Then
'            ObjLISComCode.DeptCd.KeyChange txtWardId.Text
'            lblWardNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'        Else
'            txtWardId.Text = ""
'            lblWardNm.Caption = ""
'        End If
'    Else
'        If ObjLISComCode.WardID.Exists(txtWardId.Text) = True Then
'            ObjLISComCode.WardID.KeyChange txtWardId.Text
'            lblWardNm.Caption = ObjLISComCode.WardID.Fields("wardnm")
'        Else
'            txtWardId.Text = ""
'            lblWardNm.Caption = ""
'        End If
'    End If
End Sub


Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    
    Dim objPatient As New clsPatient     '환자 클래스
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    End If
    
    If MsgFg Then Exit Sub
      
    If txtPtId.Text = "" Then
        'txtPtId.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
'        If Trim(txtPtId.Text) <> "" And .PtntQuery(txtPtId.Text) Then
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSexAge.Caption = .SEXNM & " / " & .Age & " " & .AGEDIV
            If .WardID = "" Then
                lblWard.Caption = ""
            Else
                lblWard.Caption = .WardID & "-" & .ROOMID
            End If
            PtFg = True
            ClearFg = False
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbInformation
            
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Set objPatient = Nothing
            Exit Sub
        End If
    End With
    
    Set objPatient = Nothing

    Exit Sub

End Sub

Private Function DisplayOrders(ByVal pPtId As String, ByVal pPtNm As String, ByVal pVfyDt As String, ByVal pTestDiv As String) As Boolean

    Dim i As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New Recordset
    Dim SvKeyDt As String, SvSpcNm As String
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    Dim strKeyFld As String
    Dim MySql As New clsLISSqlReview     'Sql문 클래스
    Dim tVfyDt As String
    
    Dim strNotice As String
    Dim strTmp As String
   
    'barStatus.Value = (pTestDiv + 1) * 30
    'lblStatus.Caption = lblPtNm.Caption & " 님의 " & Choose(pTestDiv + 1, "일반", "특수", "미생물") & "검사 결과내역을 검색중입니다..."
   
    Me.Enabled = False
   
    MouseRunning
    
    '** 변경 : 조회일자 범위설정에 따른 수정 By M.G.Choi
    If optPrint(2).Value = True Or optPrint(3).Value = True Then
        tVfyDt = Format(dtpVfyDt.Value, CS_DateDbFormat)
    Else
        tVfyDt = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    End If
    
    '** 원본 ----------------------------------------------------------
'    tVfyDt = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    '------------------------------------------------------------------

    '처방일/접수일 기준
    SqlStmt = MySql.SqlQueryAllResults(pPtId, "examdt", tVfyDt, pVfyDt, pTestDiv)
    
    'Query
    tmpRs.Open SqlStmt, DBConn
    
    SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    ReDim aryMesg(0)
    DisplayOrders = False
    
    If tmpRs.EOF Then GoTo Nodata
    
    With tblOrdSheet
      
        '.ReDraw = False
      
        Do Until tmpRs.EOF
         
            If Trim("" & tmpRs.Fields("RstCd").Value) = "" Then GoTo Skip
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Value = pPtId: .ForeColor = DCM_Gray
            .Col = 2: .Value = pPtNm:  .ForeColor = DCM_Gray
            
            If SvKeyDt <> Trim("" & tmpRs.Fields("KeyDate").Value) Then
                .Col = 3:   .Value = Trim("" & tmpRs.Fields("KeyDate").Value)
                            .FontBold = True: .ForeColor = vbBlack       '-- 보고일
                .Col = 4:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                            .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                SvKeyDt = Trim("" & tmpRs.Fields("KeyDate").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                .Col = 1:   .FontBold = True: .ForeColor = vbBlack
                .Col = 2:   .FontBold = True: .ForeColor = vbBlack
            Else
                .Col = 3:   .Value = "":
                            .FontBold = True: .ForeColor = vbBlack       '-- 처방일
                            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                                .Col = 4:
                                .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                            Else
                                .Col = 4:
                                .Value = "":
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                            End If
            End If
            
            .Col = 34:  .Value = Trim("" & tmpRs.Fields("KeyDate").Value)    '처방일
            .Col = 35:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
            
            .Col = 5:   '-- 검사명
                        .ForeColor = DCM_MidBlue
                        tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            
                            .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                        Else
                            .Value = Space(4) & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")
                        End If
                        
            .Col = 6:   '-- 결과명(코드일 경우..)
                        .ForeColor = DCM_Brown   '갈색
                        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
                            .Value = "미확"
                            .ForeColor = DCM_MidGray: .FontBold = False:
                        Else
                            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                                .TypeHAlign = TypeHAlignCenter
                                .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                            Else
                                .CellType = CellTypeEdit
                                .TypeHAlign = TypeHAlignLeft
                                .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                            End If
                            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                                .Value = "Growth"
                            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                                .Value = Space(3)
                            End If
                        End If
                        
            .Col = 7:   '-- 결과단위
                        .Value = Trim("" & tmpRs.Fields("RstUnit").Value)
            
            .Col = 8    '-- High / Low
                        .Value = ""
                        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG:   .ForeColor = DCM_LightBlue
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                        End If
            
            .Col = 9:   '-- Delta/Panic
                        .Value = Trim("" & tmpRs.Fields("DPDiv").Value): .ForeColor = vbRed
            
            .Col = 10:   '-- 참고치
                        If Trim("" & tmpRs.Fields("RstDiv").Value) <> "*" And Trim("" & tmpRs.Fields("TestDiv").Value) < "4" Then .Value = CS_QuestionMark
            
            .Col = 11:   '-- More Result...
                        .Value = "": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.Fields("TxtFg").Value) > "0" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("TxtFg").Value) = "Y" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then .Value = CS_FingerMark
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            If Trim("" & tmpRs.Fields("FootNoteFg").Value) = "1" Then .Value = CS_FingerMark
                            If Trim("" & tmpRs.Fields("RmkCd").Value) <> "" Then .Value = CS_FingerMark
                        End If
                        If Trim("" & tmpRs.Fields("DcFg").Value) = "1" Then .Value = .Value & "*"
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "4" Then .Value = CS_FingerMark     '해부병리
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "5" Then .Value = CS_FingerMark     '혈액은행
         
            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdDate").Value)         '-- 처방일
            .Col = 13: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)           '-- 처방번호
            .Col = 14: .Value = Trim("" & tmpRs.Fields("OrdDoct").Value)        '-- 처방의
            .Col = 15: .Value = Trim("" & tmpRs.Fields("ColDtTm").Value)        '-- 채혈일시
            .Col = 16: .Value = Trim("" & tmpRs.Fields("ColId").Value)          '-- 채혈자
            .Col = 17: .Value = Trim("" & tmpRs.Fields("RcvDtTm").Value)        '-- 접수일시
            .Col = 18: .Value = Trim("" & tmpRs.Fields("RcvId").Value)          '-- 접수자
            .Col = 19: .Value = Trim("" & tmpRs.Fields("WorkArea").Value):  pWorkArea = .Value  'WorkArea
            .Col = 20: .Value = Trim("" & tmpRs.Fields("AccDt").Value):     pAccDt = .Value     'AccDt
            .Col = 21: .Value = Trim("" & tmpRs.Fields("AccSeq").Value):    pAccSeq = .Value    'AccSeq
            .Col = 22: .Value = Trim("" & tmpRs.Fields("LastRst").Value)        '-- 최근결과
            .Col = 23: .Value = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)     '-- 최근결과일시
            .Col = 24: .Value = Trim("" & tmpRs.Fields("LastVfyId").Value)      '-- 최근결과 보고자
            .Col = 25: .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)        '-- 보고일시
            .Col = 26: .Value = Trim("" & tmpRs.Fields("VfyId").Value)          '-- 보고자
            .Col = 27: .Value = Trim("" & tmpRs.Fields("Sex").Value)            '-- Sex
            .Col = 28: .Value = Trim("" & tmpRs.Fields("AgeDay").Value)         '-- AgeDay
            .Col = 29: .Value = Trim("" & tmpRs.Fields("TestCd").Value)         '-- 검사코드
            .Col = 30: .Value = Trim("" & tmpRs.Fields("SpcCd").Value)          '-- 검체코드
            .Col = 31: .Value = Trim("" & tmpRs.Fields("VfyDt").Value)          '-- 보고일
            .Col = 32: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)        '-- 검사구분
            .Col = 33: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)         '-- 진료과
            .Col = 36: .Value = Trim("" & tmpRs.Fields("TxtFg").Value)          '-- 소견결과여부
            .Col = 37: .Value = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
            .Col = 38: .Value = Trim("" & tmpRs.Fields("RmkCd").Value)          '-- Remark 코드
            .Col = 39: .Value = Trim("" & tmpRs.Fields("SenFg").Value)          '-- 감수성 여부
            .Col = 40: .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)         '-- 처방구분
            .Col = 41: .Value = Trim("" & tmpRs.Fields("UnitQty").Value)        '-- 수혈수량
            .Col = 42: .Value = Trim("" & tmpRs.Fields("ReqDt").Value)          '-- 수혈예정일
            .Col = 43: .Value = Trim("" & tmpRs.Fields("ReqTm").Value)          '-- 수혈예정시간
            .Col = 44: .Value = Trim("" & tmpRs.Fields("WardId").Value)         '-- 병동
            .Col = 45: .Value = Trim("" & tmpRs.Fields("HosilId").Value)        '-- 호실
            .Col = 46: .Value = Trim("" & tmpRs.Fields("RoomId").Value)        '-- 호실
            
'            ReDim Preserve aryMesg(UBound(aryMesg) + 1)
'            aryMesg(UBound(aryMesg)) = Trim("" & tmpRs.Fields("Mesg"))    '-- 진료과Remark
            If Trim("" & tmpRs.Fields("Notice").Value) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5
                .TypeEditMultiLine = False
                .ForeColor = vbBlack
                .Value = "☞ Clinical Notice "  '& vbCrLf & Trim("" & tmpRs.Fields("Notice"))
                .RowHeight(.MaxRows) = .MaxTextRowHeight(.MaxRows)
                strNotice = Trim("" & tmpRs.Fields("Notice").Value)
                strNotice = Replace(strNotice, vbCr, "")
                strTmp = medShift(strNotice, vbLf)
                While strTmp <> ""
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 5
                    .TypeEditMultiLine = False
                    .ForeColor = &H747474
                    .Value = strTmp
                    strTmp = medShift(strNotice, vbLf)
                Wend
            End If
      
         
            DisplayOrders = True
Skip:
            tmpRs.MoveNext
        Loop
      
        .Row = -1: .Col = 6: .Col2 = 5
        .BlockMode = True
        .AllowCellOverflow = True
        .BlockMode = False
      
        .RowHeight(-1) = 11.5
        .ReDraw = True
      
        'If chkRefVal.Value = 0 Then GoTo ExitPos
        GoTo ExitPos
      
        Dim tmpTestCd As String
        Dim tmpSpcCd As String
        Dim tmpVfyDt As String
        Dim tmpSex As String
        Dim tmpAgeDay As String
        Dim tmpRs1 As New Recordset
        Dim tmpRefFromVal As Double
        Dim tmpRefToVal As Double
        Dim tmpRefCd As String
      
        DoEvents
        For i = 1 To .MaxRows
            '참고치 검색
            .Row = i
            .Col = 10: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 27:  tmpSex = Trim(.Value)
            .Col = 28:  tmpAgeDay = Trim(.Value)
            .Col = 29:  tmpTestCd = Trim(.Value)
            .Col = 30:  tmpSpcCd = Trim(.Value)
            .Col = 31:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then
                '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
                '--> 거의 Both로 등록됨.
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open SqlStmt, DBConn
                
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            Set tmpRs1 = Nothing
            For j = i To .MaxRows
                .Row = j
                .Col = 29   '참고치
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 10: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents

RefSkip:
        Next
      
ExitPos:
        'If .MaxRows < 20 Then .MaxRows = 20
      
    End With
   
Nodata:
    Me.Enabled = True
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Function
