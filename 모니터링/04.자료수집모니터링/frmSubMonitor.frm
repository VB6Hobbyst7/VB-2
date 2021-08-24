VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmSubMonitor 
   Caption         =   "한국해양조사협회 자료수집 현황 모니터링 (상세조회)"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16590
   Icon            =   "frmSubMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   16590
   StartUpPosition =   1  '소유자 가운데
   Begin VB.ComboBox cmbxSechDTEdHour 
      Height          =   300
      IMEMode         =   1  '입력 상태 설정
      Left            =   6885
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   150
      Width           =   615
   End
   Begin VB.ComboBox cmbxSechDTStHour 
      Height          =   300
      Left            =   4785
      Style           =   2  '드롭다운 목록
      TabIndex        =   5
      Top             =   165
      Width           =   615
   End
   Begin VB.CommandButton btnDtSearch 
      Caption         =   "검색"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7575
      TabIndex        =   4
      Top             =   165
      Width           =   735
   End
   Begin VB.TextBox txtSechDTEdDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   5670
      MaxLength       =   10
      TabIndex        =   3
      Top             =   165
      Width           =   1170
   End
   Begin VB.TextBox txtSechDTStDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   3525
      MaxLength       =   10
      TabIndex        =   2
      Top             =   165
      Width           =   1215
   End
   Begin VB.ComboBox cmbxSechDTID 
      Height          =   300
      ItemData        =   "frmSubMonitor.frx":058A
      Left            =   1170
      List            =   "frmSubMonitor.frx":058C
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   150
      Width           =   1575
   End
   Begin FPSpreadADO.fpSpread spdDetail 
      Height          =   9045
      Left            =   330
      TabIndex        =   0
      Top             =   570
      Width           =   15885
      _Version        =   458752
      _ExtentX        =   28019
      _ExtentY        =   15954
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
      SpreadDesigner  =   "frmSubMonitor.frx":058E
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5445
      TabIndex        =   9
      Top             =   165
      Width           =   165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "관측소 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   8
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "기간 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   7
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmSubMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fpSpread_Tot_DtVPN_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub Form_Load()
    
    Call Init_fpSpread_Detail

End Sub
