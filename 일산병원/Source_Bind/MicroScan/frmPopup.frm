VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPopup 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   1515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtQry 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1230
      Width           =   1755
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1395
      _Version        =   393216
      _ExtentX        =   2461
      _ExtentY        =   5847
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmPopup.frx":0000
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_orgnm As ADODB.Recordset


Private Sub Form_Load()

    Dim rs_orgnm As New ADODB.Recordset


    Set rs_orgnm = cn_Ser.Execute(txtQry.Text)
    Do Until rs_orgnm.EOF
        'strOrgNm = rs_orgnm.Fields(0).Value & ""
'        SetText vasCode, strOrgNm, lngRow, 3
        rs_orgnm.MoveNext
    Loop
            
    Set rs_orgnm = Nothing


End Sub
