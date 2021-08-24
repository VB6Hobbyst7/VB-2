VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS308 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Blood History"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBBS308.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   14535
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin VB.CheckBox chkBar 
      BackColor       =   &H00800000&
      Caption         =   "¹ÙÄÚµå·Î Ã³¸®"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1500
      TabIndex        =   20
      Top             =   150
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Ãâ·Â(&P)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¶È¸(&Q)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12465
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   1500
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "È­¸éÁö¿ò(&C)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   2
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¾·á(&X)"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   3
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblHistory 
      Height          =   6150
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   10848
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿òÃ¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   9
      MaxRows         =   23
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS308.frx":076A
      UserResize      =   0
      TextTip         =   2
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   360
      Left            =   75
      TabIndex        =   21
      Top             =   45
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   635
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Ç÷¾× History"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   75
      TabIndex        =   5
      Top             =   345
      Width           =   14370
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   5700
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ç÷¾×Çü"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   5700
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   5700
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   975
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ã¤Ç÷ÀÚ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   17
         Left            =   75
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ç÷¾×Á¦Á¦"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   75
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ç÷¾×¹øÈ£"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   19
         Left            =   75
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   975
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ÁöÁ¤Ç÷¾×¿©ºÎ"
         Appearance      =   0
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1290
         MaxLength       =   13
         TabIndex        =   17
         Top             =   225
         Width           =   1845
      End
      Begin VB.ComboBox cboComp 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBBS308.frx":11B6
         Left            =   3150
         List            =   "frmBBS308.frx":11B8
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   16
         Top             =   225
         Width           =   2535
      End
      Begin MedControls1.LisLabel lblCompo 
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtid 
         Height          =   330
         Left            =   1890
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   975
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColNm 
         Height          =   315
         Left            =   6915
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   975
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColDt 
         Height          =   315
         Left            =   9525
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   975
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTransform 
         Height          =   330
         Left            =   1290
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   975
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblavailable 
         Height          =   315
         Left            =   6915
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVolume 
         Height          =   315
         Left            =   9525
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColTm 
         Height          =   315
         Left            =   9525
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   315
         Left            =   6915
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   2970
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   975
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCenterNm 
         Height          =   315
         Left            =   7620
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCenterCd 
         Height          =   315
         Left            =   6915
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   8325
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "¿ë·®"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   8325
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   975
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ã¤Ç÷ÀÏÀÚ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   5700
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "À¯È¿±â°£"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   8325
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Ã¤Ç÷½Ã°£"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS308"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn
    Dt = 1
    Contents
    Charge
    Method
    SpcNum
    Nurse
    Ptid
    ptnm
    Etc
End Enum

Private Sub cboComp_Click()
    
    If txtBldNo = "" Then Exit Sub
    If cboComp.ListIndex < 0 Then Exit Sub
    
    Dim objGetSql  As New clsGetSqlStatement
    Dim RS       As New Recordset
    Dim strCompocd As String
    Dim strBldNo   As String
    
    Set objGetSql = New clsGetSqlStatement
    Set RS = New Recordset
    
    strCompocd = medGetP(cboComp.List(cboComp.ListIndex), 2, COL_DIV)
    If chkBar.value = 1 Then
        strBldNo = Mid(txtBldNo, 1, 2) & "-" & Mid(txtBldNo, 3, 2) & "-" & Mid(txtBldNo, 5, 6)
    Else
        strBldNo = txtBldNo
    End If
    
    Set RS = objGetSql.Get_BLOOD_INFORMATION(strBldNo, strCompocd)
    
    If Not RS.EOF Then
        With RS
            lblCompo.Caption = Trim(medGetP(cboComp.List(cboComp.ListIndex), 1, COL_DIV))
            lblCenterCd.Caption = medGetP(cboComp.List(cboComp.ListIndex), 3, COL_DIV)
            lblCenterNm.Caption = medGetP(cboComp.List(cboComp.ListIndex), 4, COL_DIV)
            
            lblABO.Caption = .Fields("abo").value & "" & .Fields("rh").value & ""
            lblColDt.Caption = Format(.Fields("coldt").value & "", "####-##-##")
            lblColTm.Caption = Format(.Fields("coltm").value & "", "##:##")
            lblAvailable.Caption = .Fields("available").value & ""
            lblVolume.Caption = .Fields("volumn").value & ""
        
            If .Fields("reserved").value & "" <> "1" And .Fields("autofg").value & "" <> "1" Then
                lblTransform.Caption = ""
                lblPtId.Caption = ""
                lblPtNm.Caption = ""
            End If
            If .Fields("reserved").value & "" <> "0" Then
                lblTransform.Caption = "ÁöÁ¤"
                lblPtId.Caption = .Fields("ptid").value & ""
                lblPtNm.Caption = GetPtNm(.Fields("ptid").value & "")
            End If
            If .Fields("autofg").value & "" <> "0" Then
                lblTransform.Caption = "ÀÚ°¡"
                lblPtId.Caption = .Fields("ptid").value & ""
                lblPtNm.Caption = GetPtNm(.Fields("ptid").value & "")
            End If
            

            lblColNm.Caption = GetEmpNm(.Fields("colid").value & "")
            
        End With
    End If
    Call ICSPatientMark(lblPtId.Caption, enICSNum.BBS_ALL)
    Set RS = Nothing
    Set objGetSql = Nothing
End Sub

Private Sub chkBar_Click()
    txtBldNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    Clear
    txtBldNo = ""
    txtBldNo.SetFocus
'    chkBar.value = 1
End Sub
Sub Clear()
    lblPtId.Caption = ""
    lblCompo.Caption = ""
    lblAvailable.Caption = ""
    lblColNm.Caption = ""
    lblColDt.Caption = ""
    lblColTm.Caption = ""
    lblABO.Caption = ""
    lblVolume.Caption = ""
    lblTransform.Caption = ""
    lblCenterCd.Caption = ""
    lblCenterNm.Caption = ""
    
    'cboComp.Clear
    'medClearTable tblHistory
    tblHistory.MaxRows = 0
    Call ICSPatientMark
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
'Ç÷¾×¹øÈ£º° History Á¶È¸
'Ç÷¾×¹øÈ£¸¦ ÀÔ·Â ¹ÞÀ¸¸é bbs401¿¡¼­ ±×¿¡ ÇØ´çµÇ´Â Component¸¦ ÄÞº¸¹Ú½º¿¡ º¸¿©ÁØ´Ù.
    Dim objGetSql  As clsGetSqlStatement
    Dim strCompocd As String
    Dim lngRowcnt  As Long
    Dim strBldNo   As String
    
    
    If txtBldNo = "" Then Exit Sub
    If cboComp.ListIndex < 0 Then Exit Sub
    
    Set objGetSql = New clsGetSqlStatement
    
    Me.MousePointer = 11
    
    'tblHistory.MaxRows = 0
    
    strCompocd = medGetP(cboComp.List(cboComp.ListIndex), 2, COL_DIV)
    
    If chkBar.value = 1 Then
        strBldNo = Mid(txtBldNo, 1, 2) & "-" & Mid(txtBldNo, 3, 2) & "-" & Mid(txtBldNo, 5, 6)
    Else
        strBldNo = txtBldNo
    End If
    
    Call objGetSql.Get_Ent_InFo(strBldNo, strCompocd)
    
    Call History_Display(objGetSql.BldDic)
    
    Set objGetSql = Nothing
    
    Me.MousePointer = 0
End Sub

Private Sub History_Display(objDIC As clsDictionary)
    Dim ii As Integer
    
    If objDIC.RecordCount = 0 Then
        MsgBox "Ç÷¾× History ³»¿ªÀÌ ¾ø½À´Ï´Ù.", vbExclamation
        Exit Sub
    End If
    
    objDIC.Sort = True
    With tblHistory
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = objDIC.RecordCount
        objDIC.MoveFirst
            Do Until objDIC.EOF
            ii = ii + 1
            .Row = ii
            .Col = TblColumn.Dt:       .value = Format(objDIC.Fields("indexdt"), "0###-##-##") & " " & Format(Mid(objDIC.Fields("indextm"), 1, 6), "0#:##:##")
            .Col = TblColumn.Contents: .value = objDIC.Fields("contents")
            .Col = TblColumn.Charge:   .value = objDIC.Fields("chargenm")
            .Col = TblColumn.Method:   .value = objDIC.Fields("xmethod")
            .Col = TblColumn.SpcNum:   .value = objDIC.Fields("spcno")
            .Col = TblColumn.Nurse:    .value = objDIC.Fields("nurse")
            .Col = TblColumn.Ptid:     .value = objDIC.Fields("ptid")
            .Col = TblColumn.ptnm:
                If objDIC.Fields("ptid") <> "" Then
                    .value = GetPtNm(objDIC.Fields("ptid"))
                    If .value = "0" Then .value = ""
                End If
                
             '2001-11-27¼öÁ¤:
            '±æº´¿øÀÇ °æ¿ì ¾ÆÀÌµð°¡ 0ÀÌ¸é Á¦¿Ü
            If Trim(.value) <> "" And Trim(.value) <> "0" Then
                .Col = TblColumn.ptnm: .value = GetPtNm(objDIC.Fields("ptid"))
            Else
                .Col = TblColumn.ptnm: .value = ""
            End If
            
            .Col = TblColumn.Etc:      .value = objDIC.Fields("etc")
            objDIC.MoveNext
        Loop
        .ReDraw = True
    End With
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    chkBar.value = 1
End Sub

Private Sub Form_Load()
    Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub tblHistory_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim vTipText As Variant
    
    With tblHistory
        If .DataRowCnt = 0 Then Exit Sub
        If Col <> 9 Then Exit Sub
        If Row < 1 Or Row > .DataRowCnt Then Exit Sub
        
        Call .GetText(Col, Row, vTipText)
        If vTipText = "" Then Exit Sub
        
        Call .SetTextTipAppearance("±¼¸²Ã¼", 9, False, False, &HEEFDF2, vbBlack)
        
        TipText = vbNewLine & _
                  Space(5) & "¡Ø ºñ °í ¡Ø" & _
                  vbNewLine & _
                  vbNewLine & _
                  Space(5) & "- " & Replace(vTipText, ",", vbNewLine & Space(5) & "- ") & _
                  vbNewLine
        TipWidth = 5000
        MultiLine = 1
        ShowTip = True
    End With
End Sub

Private Sub txtBldNo_Change()
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    txtBldNo.tag = Trim(txtBldNo)
    txtBldNo.SelStart = 0
    txtBldNo.SelLength = Len(txtBldNo)
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtBldNo_LostFocus()
'Ç÷¾×¹øÈ£º° History Á¶È¸
'Ç÷¾×¹øÈ£¸¦ ÀÔ·Â ¹ÞÀ¸¸é bbs401¿¡¼­ ±×¿¡ ÇØ´çµÇ´Â Component¸¦ ÄÞº¸¹Ú½º¿¡ º¸¿©ÁØ´Ù.
    If txtBldNo = "" Then Exit Sub
    If txtBldNo = txtBldNo.tag Then Exit Sub
    
    Call QueryBlood
End Sub
Private Sub QueryBlood()
    Dim DrRS      As Recordset
    Dim objGetSql As New clsGetSqlStatement
    Dim strBldNo  As String
    
    
    If chkBar.value = 1 Then
        If Len(txtBldNo) < 7 Then Exit Sub
        strBldNo = Mid(txtBldNo, 1, 2) & "-" & _
                 Mid(txtBldNo, 3, 2) & "-" & _
                 Mid(txtBldNo, 5, 6)
    Else
        strBldNo = txtBldNo.Text
    End If
    
    Call Clear
    cboComp.Clear
    Set DrRS = objGetSql.Get_BLOOD_COMPONENT(strBldNo)
    If DrRS.EOF Then
        MsgBox "Ç÷¾×¹øÈ£¿¡ ÇØ´çµÇ´Â µ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù.", vbInformation + vbOKOnly, "Ç÷¾×History"
        Clear
        
    Else
        With DrRS
            Do Until .EOF
                cboComp.AddItem .Fields("field1").value & "" & Space(30) & COL_DIV & _
                                .Fields("compocd").value & "" & COL_DIV & _
                                .Fields("centercd").value & "" & COL_DIV & _
                                .Fields("buildnm").value & ""
                                
                .MoveNext
            Loop
            cboComp.ListIndex = 0
            cmdQuery.SetFocus
        End With
    End If
    txtBldNo.tag = txtBldNo
    Set objGetSql = Nothing
    
End Sub
