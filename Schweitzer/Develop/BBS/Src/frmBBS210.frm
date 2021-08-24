VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS210 
   BackColor       =   &H00E0E0E0&
   Caption         =   "혈액부작용등록"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9660
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtAbo 
      Height          =   375
      Left            =   5820
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1605
      Width           =   3675
   End
   Begin VB.TextBox txtColdt 
      Height          =   375
      Left            =   5820
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2010
      Width           =   3675
   End
   Begin VB.TextBox txtAva 
      Height          =   375
      Left            =   5820
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2430
      Width           =   3675
   End
   Begin VB.TextBox txtExpdt 
      Height          =   375
      Left            =   5820
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2835
      Width           =   3675
   End
   Begin VB.TextBox txtDeliveryDt 
      Height          =   375
      Left            =   5820
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3270
   End
   Begin VB.TextBox txtDeliveryNm 
      Height          =   375
      Left            =   5820
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3660
      Width           =   3675
   End
   Begin VB.TextBox txtDExpDt 
      Height          =   375
      Left            =   5820
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4065
      Width           =   3675
   End
   Begin VB.TextBox txtExpNm 
      Height          =   375
      Left            =   5820
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4485
      Width           =   3675
   End
   Begin VB.TextBox txtReqNm 
      Height          =   375
      Left            =   5820
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4890
      Width           =   3675
   End
   Begin VB.TextBox txtReason 
      Height          =   375
      Left            =   5820
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5295
      Width           =   3675
   End
   Begin VB.TextBox txtRemark 
      Height          =   1035
      Left            =   5820
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5715
      Width           =   3675
   End
   Begin VB.TextBox txtCompNm 
      Height          =   375
      Left            =   5835
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1185
      Width           =   3675
   End
   Begin VB.TextBox txtbldno 
      Height          =   360
      Left            =   5835
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   750
      Width           =   3690
   End
   Begin VB.CommandButton cmdBldNo 
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
      Height          =   350
      Left            =   6870
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   390
      Width           =   350
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   4
      Left            =   4755
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   750
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "혈액번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   6
      Left            =   4755
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "채혈일"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   3
      Left            =   4755
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1605
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "혈액형"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   7
      Left            =   4755
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1185
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "혈액제제"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   9
      Left            =   4755
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "유효일"
      Appearance      =   0
   End
   Begin VB.TextBox txtSRemark 
      Height          =   900
      Left            =   1005
      TabIndex        =   10
      Top             =   6735
      Width           =   8490
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   8190
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   7770
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   6885
      Style           =   1  '그래픽
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   7770
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   5565
      Style           =   1  '그래픽
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "15101"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '없음
      Height          =   390
      Left            =   75
      ScaleHeight     =   390
      ScaleWidth      =   4650
      TabIndex        =   1
      Top             =   360
      Width           =   4650
      Begin DRcontrol1.DrLabel lblPtNm 
         Height          =   375
         Left            =   2475
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   661
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelStyle      =   1
         Caption         =   ""
      End
      Begin DRcontrol1.DrText txtPtid 
         Height          =   375
         Left            =   1065
         TabIndex        =   2
         Top             =   15
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Appearance      =   1
         Alignment       =   2
         BorderColor     =   4210752
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   0
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   15
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자정보"
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4755
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액정보"
   End
   Begin FPSpread.vaSpread tblBloodList 
      Height          =   5970
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   4650
      _Version        =   196608
      _ExtentX        =   8202
      _ExtentY        =   10530
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
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
      MaxCols         =   10
      MaxRows         =   28
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS210.frx":0000
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   900
      Index           =   0
      Left            =   75
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6735
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1588
      BackColor       =   10392451
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Remark"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   1
      Left            =   4755
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2835
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "폐기일"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   2
      Left            =   4755
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4065
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "출고폐기일"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   8
      Left            =   4755
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3660
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "출고자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   10
      Left            =   4755
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "출고일자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   11
      Left            =   4755
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4485
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "폐기자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   12
      Left            =   4755
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5295
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "폐기사유"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   13
      Left            =   4755
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4890
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "폐기요청자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   1005
      Index           =   14
      Left            =   4755
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1773
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
      Caption         =   "폐기MEMO"
      Appearance      =   0
   End
   Begin DRcontrol1.DrText txtReasonCd 
      Height          =   375
      Left            =   5820
      TabIndex        =   26
      Top             =   375
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Appearance      =   1
      Alignment       =   2
      BorderColor     =   4210752
   End
   Begin DRcontrol1.DrLabel lblReasonNm 
      Height          =   345
      Left            =   7215
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   390
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelStyle      =   1
      Caption         =   ""
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   15
      Left            =   4755
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   375
      Width           =   1050
      _ExtentX        =   1852
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
      Caption         =   "사유"
      Appearance      =   0
   End
   Begin DRcontrol1.DrText txtNum 
      Height          =   375
      Left            =   9105
      TabIndex        =   31
      Top             =   3240
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Appearance      =   1
      Alignment       =   2
      BorderColor     =   4210752
   End
End
Attribute VB_Name = "frmBBS210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Private objSQL As New clsHospital05
Private ObjDic As New clsDictionary
Private mvarPtId As String
Public Property Let ptid(ByVal vData As String)
    mvarPtId = vData
End Property
Private Sub ScreenClear()
    lblPtNm.Caption = ""
    lblReasonNm.Caption = ""
    txtReasonCd.Text = ""
    txtbldno.Text = ""
    txtCompNm.Text = ""
    txtAbo.Text = ""
    txtColdt.Text = ""
    txtAva.Text = ""
    txtExpdt.Text = ""
    txtDeliveryDt.Text = ""
    txtDeliveryNm.Text = ""
    txtDExpDt.Text = ""
    txtExpNm.Text = ""
    txtReqNm.Text = ""
    txtReason.Text = ""
    txtRemark.Text = ""
    txtNum.Text = ""
    txtSRemark.Text = ""
    Call medClearTable(tblBloodList)
End Sub

Private Sub BloodClear()

    txtbldno.Text = ""
    txtCompNm.Text = ""
    txtAbo.Text = ""
    txtColdt.Text = ""
    txtAva.Text = ""
    txtExpdt.Text = ""
    txtDeliveryDt.Text = ""
    txtDeliveryNm.Text = ""
    txtDExpDt.Text = ""
    txtExpNm.Text = ""
    txtReqNm.Text = ""
    txtReason.Text = ""
    txtRemark.Text = ""
    txtNum.Text = ""
    txtSRemark.Text = ""
End Sub

Private Function PtQuery(ByVal qPtid As String)
    Dim strTmp    As String
    Dim strSexAge As String
'    Dim ii        As Integer
    Dim strLength As String
    '환자정보
'    For ii = 1 To BBS_PTID_LENGTH
'        strLength = strLength & "0"
'    Next
    
    strLength = String(BBS_PTID_LENGTH, "0")
    
    strTmp = objSQL.GetPtInformatio(Format(qPtid, strLength))
    Call ICSPatientMark(Format(qPtid, strLength), enICSNum.BBS_ALL)

    If strTmp <> "" Then
        strSexAge = SDA_String(medGetP(strTmp, 2, COL_DIV))
        txtPtid.Text = (Format(qPtid, strLength))
        Call ScreenClear
        lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV) & "   " & _
                          medGetP(strSexAge, 1, COL_DIV) & "/" & _
                          medGetP(strSexAge, 3, COL_DIV)
        
        Call BloodQuery(Format(qPtid, strLength))
    Else
        MsgBox txtPtid.Text & " 는 존재하지 않는 환자입니다.", vbInformation + vbOKOnly, "환자정보"
        Call ScreenClear
        txtPtid.Text = ""
    End If
End Function
Private Sub BloodQuery(ByVal qPtid As String)
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim ii As Long
    
    SSQL = objSQL.GetDeilveryList(qPtid)
    Set RS = New Recordset
    RS.Open SSQL, DBConn

    If Not RS.EOF Then
        
        With tblBloodList
            Do Until RS.EOF
                If .DataRowCnt > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .value = RS.Fields("bldsrc").value & "" & "-" & _
                                   RS.Fields("bldyy").value & "" & "-" & _
                                   Format(RS.Fields("bldno").value & "", "000000")
                .Col = 2: .value = RS.Fields("abo").value & ""
                .Col = 3: .value = RS.Fields("rh").value & ""
                .Col = 4: .value = RS.Fields("volumn").value & ""
                .Col = 5
                Select Case RS.Fields("stscd").value & ""
                    Case BBSBloodStatus.stsDELIVERY
                        .value = "출고"
                    Case BBSBloodStatus.stsEXPIRE
                        .value = "폐기"
                        .Col = 6: .value = Format(RS.Fields("realexpdt").value & "", "####-##-##")
                        .Col = 7: .value = "1"
                        .Col = 5: .value = IIf(RS.Fields("reactionfg").value & "" = "1", "부작용", "폐기")
                        .ForeColor = IIf(.value = "부작용", DCM_Red, vbBlack)
                End Select
                .Col = 8:  .value = RS.Fields("compocd").value & ""
                .Col = 9:  .value = Format(RS.Fields("deliverydt").value & "", "####-##-##")
                .Col = 10: .value = RS.Fields("deliveryseq").value & ""
                
                RS.MoveNext
            Loop
        End With
    Else
        MsgBox qPtid & " 에 대해서는 출고된 혈액이 없습니다.", vbInformation + vbOKOnly, " 출고혈액조회"
        Call ScreenClear
        txtPtid.Text = ""
    End If
    Set RS = Nothing
End Sub

Private Sub cmdBldNo_Click()
    Dim lngTop  As Long
    Dim lngLeft As Long
    
    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        
        lngTop = txtReasonCd.Top + 2350
        lngLeft = Me.Left + txtReasonCd.Left + 50
        .FormCaption = "수혈부작용 리스트"
        .ColumnHeaderText = "사유코드;부작용명"
        .LoadPopUp objSQL.GetReactionSQL ', lngTop, lngLeft, ObjDic
        txtReasonCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
        lblReasonNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
    End With


End Sub

Private Sub cmdClear_Click()
    Call ScreenClear
    txtPtid.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim SSQL As String
    
    If txtReasonCd.Text = "" Then
        MsgBox "부작용 사유를 입력하세요", vbInformation + vbOKOnly, "부작용등록"
        Exit Sub
    End If
    
    If txtbldno.Text = "" Then
        MsgBox "부작용등록 대상혈액을 선택하세요.", vbInformation + vbOKOnly, "부작용등록"
        Exit Sub
    End If
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    SSQL = objSQL.UpdateSQL_BBS501(medGetP(txtbldno.Text, 1, "-"), _
                                   medGetP(txtbldno.Text, 2, "-"), _
                                   medGetP(txtbldno.Text, 3, "-"), _
                                   Trim(medGetP(txtCompNm.Text, 1, " ")), _
                                   Trim(Replace(txtDeliveryDt, "-", "")), _
                                   Trim(txtNum.Text), _
                                   Trim(txtReasonCd.Text), _
                                   Trim(txtSRemark.Text))
                                 
    DBConn.Execute SSQL
    DBConn.CommitTrans
    MsgBox "부작용 등록이 완결되었습니다", vbInformation + vbOKOnly, "부작용등록"
    Call ScreenClear
    Call BloodQuery(txtPtid.Text)
    
    
    Exit Sub
    
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim SSQL As String
    Dim RS   As Recordset
    
    ObjDic.Clear
    ObjDic.FieldInialize "reactioncd", "reactionNm"
    
    Set RS = New Recordset
    RS.Open objSQL.GetReactionSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            If ObjDic.Exists(RS.Fields("cdval1").value & "") Then
                ObjDic.KeyChange RS.Fields("cdval1").value & ""
                ObjDic.Fields("reactionnm") = RS.Fields("field1").value & ""
            Else
                ObjDic.AddNew RS.Fields("cdval1").value & "", RS.Fields("field1").value & ""
            End If
            RS.MoveNext
        Loop
    End If
    Call ScreenClear
    txtPtid.Text = ""
    If mvarPtId <> "" Then PtQuery (mvarPtId)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mvarPtId = ""
    Call ICSPatientMark
    Set ObjDic = Nothing
    Set objSQL = Nothing
    Set objCodeList = Nothing
End Sub

Private Sub tblBloodList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim sBldNo      As String
    Dim sBldSrc     As String
    Dim sBldYY      As String
    Dim sCompocd    As String
    Dim sDeliverydt As String
    Dim sDeliverySeq As String
    
    If Row < 1 Then Exit Sub
    
    With tblBloodList
        If Row > .DataRowCnt Then Exit Sub
        Call BloodClear
        .Row = Row
        .Col = 7
        If .value = "" Then
            MsgBox "폐기되지 않은혈액은 부작용등록을 하실수 없습니다.", vbInformation + vbOKOnly, "부작용등록"
            Exit Sub
        End If
        .Col = 5
        If .value = "부작용" Then
            MsgBox "이미 부작용등록된 혈액입니다.", vbInformation + vbOKOnly, "부작용등록"
            Exit Sub
        End If
        
        .Col = 1:  sBldSrc = medGetP(.value, 1, "-")
                   sBldYY = medGetP(.value, 2, "-")
                   sBldNo = medGetP(.value, 3, "-")
        txtbldno.Text = .value
        .Col = 8:  sCompocd = .value
        .Col = 9:  sDeliverydt = .value:
        .Col = 10: sDeliverySeq = .value:
        Set RS = objSQL.GetBloodInfo(sBldSrc, sBldYY, sBldNo, sCompocd)
        If Not RS.EOF Then
            txtCompNm.Text = sCompocd & "  " & medGetP(Get_CompNm(sCompocd), 1, COL_DIV)
            txtAbo.Text = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            txtColdt.Text = Format(RS.Fields("coldt").value & "", "####-##-##")
            txtAva.Text = RS.Fields("available").value & "" & "일"
            txtExpdt.Text = Format(RS.Fields("expdt").value & "", "####-##-##")
            txtDeliveryNm.Text = objSQL.GetDeliveryNm(sBldSrc, sBldYY, sBldNo, sCompocd, Replace(sDeliverydt, "-", ""), sDeliverySeq)
            txtDExpDt.Text = Format(RS.Fields("realexpdt").value & "", "####-##-##")
            txtExpNm.Text = GetEmpNm(RS.Fields("expid").value & "")
            txtReqNm.Text = GetEmpNm(RS.Fields("exprcvid").value & "")
            txtReason.Text = RS.Fields("exprsncd").value & ""
            txtRemark.Text = RS.Fields("exprsnrmk").value & ""
            txtDeliveryDt.Text = sDeliverydt
            txtNum.Text = sDeliverySeq
        End If
    End With
    
End Sub
Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ScreenClear
        Call PtQuery(txtPtid.Text)
    End If
End Sub

Private Sub txtReasonCd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If ObjDic.Exists(txtReasonCd.Text) Then
            ObjDic.KeyChange txtReasonCd.Text
            lblReasonNm.Caption = ObjDic.Fields("reactionnm")
        End If
    End If
End Sub
