VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows 기본값
   Begin VB.OptionButton optGubun 
      BackColor       =   &H00E0E0E0&
      Caption         =   "진료"
      Height          =   225
      Index           =   0
      Left            =   8430
      TabIndex        =   4
      Top             =   1200
      Value           =   -1  'True
      Width           =   765
   End
   Begin VB.OptionButton optGubun 
      BackColor       =   &H00E0E0E0&
      Caption         =   "검진"
      Height          =   225
      Index           =   1
      Left            =   9375
      TabIndex        =   3
      Top             =   1200
      Width           =   765
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   3975
      Left            =   270
      TabIndex        =   2
      Top             =   2100
      Width           =   10245
      _Version        =   393216
      _ExtentX        =   18071
      _ExtentY        =   7011
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
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   8370
      TabIndex        =   1
      Top             =   390
      Width           =   1785
   End
   Begin VB.TextBox txtSQL 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   300
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If optGubun(0).Value = True Then
        res = db_select_Vas(gServer, txtSQL, vasList)
    Else
        res = db_select_Vas(gServer1, txtSQL, vasList)
    End If
End Sub
