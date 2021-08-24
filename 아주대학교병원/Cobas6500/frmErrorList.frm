VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmErrorList 
   Caption         =   "미전송 결과 리스트"
   ClientHeight    =   7590
   ClientLeft      =   6660
   ClientTop       =   2595
   ClientWidth     =   6555
   LinkTopic       =   "Form3"
   ScaleHeight     =   7590
   ScaleWidth      =   6555
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   30
      Width           =   135
   End
   Begin FPSpread.vaSpread vasErrorList 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   13150
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmErrorList.frx":0000
   End
End
Attribute VB_Name = "frmErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    SQL = ""
    SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
    res = SendQuery(gLocal, SQL)
End Sub
