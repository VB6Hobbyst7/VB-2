VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmPart 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검사SLip별 통계"
   ClientHeight    =   5685
   ClientLeft      =   540
   ClientTop       =   1305
   ClientWidth     =   11145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprPart 
      Height          =   4290
      Left            =   180
      TabIndex        =   3
      Top             =   1035
      Width           =   6675
      _Version        =   196608
      _ExtentX        =   11774
      _ExtentY        =   7567
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   30
      MaxRows         =   100
      SpreadDesigner  =   "frmPart.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   600
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   6720
      _Version        =   65536
      _ExtentX        =   11853
      _ExtentY        =   1058
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36446
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   420
         Left            =   3240
         TabIndex        =   4
         Top             =   90
         Width           =   1680
         Caption         =   "조회확인"
         Size            =   "2963;741"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtFrDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()

    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
